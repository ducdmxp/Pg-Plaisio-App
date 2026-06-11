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

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<ModelLine> lines = SelectLines(uiDoc);
            if (lines.Count == 0)
            {
                TaskDialog.Show("Convert 2D to 3D", "No lines selected.");
                return Result.Cancelled;
            }

            List<ModelLine> tuongchinhs = GetLinesByStyle(lines, "NOSIVO");
            if (tuongchinhs.Count == 0)
            {
                TaskDialog.Show("Convert 2D to 3D", "No NOSIVO lines found.");
                return Result.Cancelled;
            }

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

            List<List<ModelLine>> groups = GroupParallelLines(tuongchinhs);

            Transaction tran = new Transaction(doc, "Convert Lines to Walls");

            try
            {
                tran.Start();
                foreach (var group in groups)
                {
                    CreateWallFromGroup(doc, group, level);
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

        private List<List<ModelLine>> GroupParallelLines(List<ModelLine> lines)
        {
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

        private void CreateWallFromGroup(Document doc, List<ModelLine> group, Level level)
        {
            ModelLine mline1 = group
                .OrderByDescending(l => l.GeometryCurve.Length)
                .FirstOrDefault();

            ModelLine mline2 = group
                .Where(l => l != null && l.Id != mline1.Id)
                .FirstOrDefault(l => Common.IsParallel((l.GeometryCurve as Line)?.Direction, (mline1.GeometryCurve as Line)?.Direction) == true
                && Common.IsCollinear((l.GeometryCurve as Line), (mline1.GeometryCurve as Line)));

            if (mline2 == null) return;

            Line line1 = mline1.GeometryCurve as Line;
            Line line2 = mline2?.GeometryCurve as Line;

            if (line2 == null) return;

            XYZ center2 = line2.Evaluate(0.5, true);
            XYZ pointProjection = Common.GetPointProjectOnLine(line1, center2);
            if (pointProjection == null) return;

            XYZ vector = (center2 - pointProjection).Normalize();

            double withWall = pointProjection.DistanceTo(center2);
            if (withWall < 1e-6) return;

            Line lineCenter = line1.CreateTransformed(Transform.CreateTranslation(vector * withWall / 2)) as Line;
            if (lineCenter == null) return;

            WallType wallType = GetOrCreateWallType(doc, withWall);
            if (wallType == null) return;

            try
            {
                Wall.Create(doc, lineCenter, wallType.Id, level.Id, 3000 / 304.8, 0, false, false);
            }
            catch { }
        }

        private WallType GetOrCreateWallType(Document doc, double thickness)
        {
            string typeName = $"Wall_{Math.Round(thickness * 304.8, 1)}mm";

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

                CompoundStructure cs = newType.GetCompoundStructure();
                if (cs != null && cs.LayerCount > 0)
                {
                    cs.SetLayerWidth(0, thickness);
                    newType.SetCompoundStructure(cs);
                }

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