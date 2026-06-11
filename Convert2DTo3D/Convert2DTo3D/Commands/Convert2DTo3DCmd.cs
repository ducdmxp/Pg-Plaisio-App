using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ParameterUtils = Convert2DTo3D.Utils.ParameterUtils;

namespace Convert2DTo3D.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class Convert2DTo3DCmd : IExternalCommand
    {
        // 23mm in feet (Revit internal unit)
        private const double WithdDefaut = 23.0 / 304.8;

        private const string layerNameExterior = "NOSIVO";
        private const string layerNameInterior1 = "POMOCNO";
        private const string layerNameInterior2 = "PREGRADE";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            Level level = new FilteredElementCollector(doc)
               .OfClass(typeof(Level))
               .Cast<Level>()
               .OrderBy(l => l.Elevation)
               .FirstOrDefault();

            if (level == null)
            {
                TaskDialog.Show("Convert 2D to 3D", "No level found.");
                return Result.Cancelled;
            }

            List<ModelLine> lines = SelectLines(uiDoc);
            if (lines.Count == 0)
            {
                TaskDialog.Show("Convert 2D to 3D", "No lines selected.");
                return Result.Cancelled;
            }

            List<ModelLine> lstMLineExterior = GetLinesByStyle(lines, layerNameExterior);

            List<ModelLine> lstMLineInterior1 = GetLinesByStyle(lines, layerNameInterior1);

            List<ModelLine> lstMLineInterior2 = GetLinesByStyle(lines, layerNameInterior2);

            List<List<ModelLine>> groupExteriors = GroupParallelLines(lstMLineExterior);
            List<List<ModelLine>> groupInterior1s = GroupParallelLines(lstMLineInterior1);
            List<List<ModelLine>> groupInterior2s = GroupParallelLines(lstMLineInterior2);

            Transaction tran = new Transaction(doc, "Convert Lines to Walls");

            try
            {
                tran.Start();
                foreach (var group in groupExteriors)
                {
                    CreateWallFromGroup(doc, group, level, 3000 / 304.8, layerNameExterior);
                }
                foreach (var group in groupInterior1s)
                {
                    CreateWallFromGroup(doc, group, level, 3000 / 304.8, layerNameInterior1);
                }
                foreach (var group in groupInterior2s)
                {
                    CreateWallFromGroup(doc, group, level, 3000 / 304.8, layerNameInterior2);
                }
                tran.Commit();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                tran.RollBack();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Nhóm các ModelLine trùng nhau (collinear), trả về list các cặp (line dài nhất của mỗi nhóm).
        /// Đầu vào: list các line đã song song với nhau.
        /// Kết quả: mỗi phần tử là 1 cặp [longestA, longestB] đại diện cho 2 nhóm collinear.
        /// </summary>
        private (Line lineA, Line lineB) GroupCollinearAndGetLongest(List<ModelLine> parallelLines)
        {
            var used = new HashSet<int>();
            var collinearGroups = new List<List<ModelLine>>();

            for (int i = 0; i < parallelLines.Count; i++)
            {
                if (used.Contains(i)) continue;

                Line lineI = parallelLines[i].GeometryCurve as Line;
                if (lineI == null) continue;

                var group = new List<ModelLine> { parallelLines[i] };
                used.Add(i);

                for (int j = i + 1; j < parallelLines.Count; j++)
                {
                    if (used.Contains(j)) continue;

                    Line lineJ = parallelLines[j].GeometryCurve as Line;
                    if (lineJ == null) continue;

                    if (Common.IsCollinear(lineI, lineJ))
                    {
                        group.Add(parallelLines[j]);
                        used.Add(j);
                    }
                }

                collinearGroups.Add(group);
            }

            // Ghép từng cặp nhóm collinear liền kề, lấy line dài nhất mỗi nhóm

            if (collinearGroups.Count != 2)
            {
                return (null, null);
            }

            Line line1 = GetLongestLine(collinearGroups[0]);
            Line line2 = GetLongestLine(collinearGroups[1]);

            if (line1.Length < line2.Length)
                (line1, line2) = (line2, line1);

            return (line1, line2);
        }

        private Line GetLongestLine(List<ModelLine> lines)
        {
            List<XYZ> lstPoint1 = new List<XYZ>();

            foreach (var item in lines)
            {
                Line line = item.GeometryCurve as Line;
                if (line == null) continue;

                lstPoint1.Add(line.GetEndPoint(0));
                lstPoint1.Add(line.GetEndPoint(1));
            }

            XYZ p0 = lstPoint1[0];
            XYZ p1 = lstPoint1[1];
            double distMin = p0.DistanceTo(p1);
            for (int i = 0; i < lstPoint1.Count; i++)
            {
                for (int j = i + 1; j < lstPoint1.Count; j++)
                {
                    double dist = lstPoint1[i].DistanceTo(lstPoint1[j]);
                    if (dist > distMin)
                    {
                        p0 = lstPoint1[i];
                        p1 = lstPoint1[j];
                        distMin = dist;
                    }
                }
            }

            return Line.CreateBound(p0, p1);
        }

        private List<List<ModelLine>> GroupParallelLines(List<ModelLine> lines)
        {
            if (lines.Count <= 0) return new List<List<ModelLine>>();

            var used = new HashSet<int>();
            var groups = new List<List<ModelLine>>();

            for (int i = 0; i < lines.Count; i++)
            {
                if (used.Contains(i)) continue;

                var group = new List<ModelLine> { lines[i] };
                used.Add(i);

                Line lineI = lines[i].GeometryCurve as Line;
                if (lineI == null) continue;

                for (int j = i + 1; j < lines.Count; j++)
                {
                    if (used.Contains(j)) continue;

                    Line lineJ = lines[j].GeometryCurve as Line;
                    if (lineJ == null) continue;

                    if (!Common.IsParallel(lineI.Direction, lineJ.Direction)) continue;

                    if (lineI.Length <= WithdDefaut || lineJ.Length <= WithdDefaut) continue;

                    XYZ centerJ = lineJ.Evaluate(0.5, true);
                    XYZ projected = Common.GetPointProjectOnLine(lineI, centerJ);
                    if (projected == null) continue;

                    double dist = projected.DistanceTo(centerJ);
                    if (dist <= WithdDefaut || Common.IsEqual(dist, WithdDefaut))
                    {
                        group.Add(lines[j]);
                        used.Add(j);
                    }
                }

                if (group.Count >= 2)
                    groups.Add(group);
            }

            return groups;
        }

        private void CreateWallFromGroup(Document doc, List<ModelLine> group, Level level, double height = 3000 / 304.8, string layerName = "Exterior")
        {
            (Line lineA, Line lineB) = GroupCollinearAndGetLongest(group);

            Line line1 = lineA;
            Line line2 = lineB;

            if (line1 == null || line2 == null) return;

            XYZ center2 = line2.Evaluate(0.5, true);
            XYZ pointProjection = Common.GetPointProjectOnLine(line1, center2);
            if (pointProjection == null) return;

            XYZ vector = (center2 - pointProjection).Normalize();

            double withWall = pointProjection.DistanceTo(center2);
            if (withWall < 1e-6) return;

            Line lineCenter = line1.CreateTransformed(Transform.CreateTranslation(vector * withWall / 2)) as Line;
            if (lineCenter == null) return;

            WallType wallType = GetOrCreateWallType(doc, withWall, layerName);
            if (wallType == null) return;

            try
            {
                Wall.Create(doc, lineCenter, wallType.Id, level.Id, height, 0, false, false);
            }
            catch { }
        }

        private WallType GetOrCreateWallType(Document doc, double thickness, string layerName = "Exterior")
        {
            string typeName = $"Wall_{layerName}_{Math.Round(thickness * 304.8, 1)}mm";

            WallType existing = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(wt => wt.Name == typeName);

            if (existing != null) return existing;

            WallType baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault();

            if (baseType == null) return null;

            try
            {
                WallType newType = baseType.Duplicate(typeName) as WallType;
                if (newType == null) return null;

                var structureLayer = new CompoundStructureLayer(thickness, MaterialFunctionAssignment.Structure, ElementId.InvalidElementId);
                CompoundStructure cs = CompoundStructure.CreateSingleLayerCompoundStructure(MaterialFunctionAssignment.Structure, thickness, ElementId.InvalidElementId);
                cs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                cs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);
                newType.SetCompoundStructure(cs);

                return newType;
            }
            catch
            {
                return baseType;
            }
        }

        public List<ModelLine> GetLinesByStyle(List<ModelLine> lines, string styleName)
        {
            return lines.Where(line => GetLineStyleName(line) == styleName).ToList();
        }

        public string GetLineStyleName(ModelLine line)
        {
            if (ParameterUtils.GetValueParameterByBuilt(line, BuiltInParameter.BUILDING_CURVE_GSTYLE) is ElementId lineStyleId)
            {
                if (lineStyleId != ElementId.InvalidElementId)
                {
                    var lineStyle = line.Document.GetElement(lineStyleId);
                    if (lineStyle != null)
                    {
                        return lineStyle.Name;
                    }
                }
            }

            return string.Empty;
        }

        public static List<ModelLine> SelectLines(UIDocument uidoc)
        {
            List<ModelLine> curves = new List<ModelLine>();

            try
            {
                curves = uidoc.Selection.PickObjects(ObjectType.Element, new DWGLineSelectionFilter(), "Select lines from the DWG :").Select(reference =>
                {
                    Element element = uidoc.Document.GetElement(reference);
                    return element;
                })
                .Cast<ModelLine>()
                .Where(line => line != null).ToList();
            }
            catch (Exception)
            {
            }

            return curves;
        }
    }
}