using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CmdCreateHatchAuto : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Wall> wallList = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(Wall)).Cast<Wall>().ToList();

            if (wallList?.Count <= 0) return Result.Cancelled;

            // 2. Get a Level (e.g., Level 1)
            Level level = doc.GetElement(wallList.FirstOrDefault().LevelId) as Level ?? new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault();

            if (level == null) return Result.Cancelled;

            FilledRegionType regionType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (regionType == null) return Result.Cancelled;

            double offset = 10000 / 304.8;

            List<XYZ> lstPoint = GetAllGeometrys(doc, wallList, out XYZ direction, out Solid solidTotal, offset);

            Transaction tran = new Transaction(doc, "Create Hatch Auto");

            try
            {
                tran.Start();

                List<ElementId> lstDeletes = CreateWallBoudarys(doc, level, lstPoint);

                Solid solid = CreateSolidBoudaryBuilding(doc, level, lstPoint, direction, ref lstDeletes, offset);

                Solid solidOutput = BooleanOperationsUtils.ExecuteBooleanOperation(solid, solidTotal, BooleanOperationsType.Difference);

                List<CurveLoop> loopList = GetProfiles(solidOutput);

                FilledRegion region = FilledRegion.Create(doc, regionType.Id, doc.ActiveView.Id, loopList);

                doc.Delete(lstDeletes);

                tran.Commit();
            }
            catch (Exception)
            {
                tran.RollBack();
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private List<CurveLoop> GetProfiles(Solid solid)
        {
            List<CurveLoop> profiles = new List<CurveLoop>();

            if (solid == null) return profiles;

            PlanarFace planarFace = Common.GetSolidFaces(solid)
                                        .Where(x => x is PlanarFace && Common.IsParallel(x.FaceNormal, XYZ.BasisZ))
                                        .OrderBy(x => x.Origin.Z)
                                        .FirstOrDefault();

            return planarFace?.GetEdgesAsCurveLoops().ToList();
        }

        private List<ElementId> CreateWallBoudarys(Document doc, Level level, List<XYZ> lstPoint)
        {
            List<ElementId> lstDeletes = new List<ElementId>();

            for (int i = 0; i < lstPoint.Count - 1; i++)
            {
                XYZ point1 = lstPoint[i];
                XYZ point2 = lstPoint[i + 1];

                Wall wall = CreateWall(doc, point1, point2, level);

                if (wall != null) lstDeletes.Add(wall.Id);
            }

            return lstDeletes;
        }

        public Wall CreateWall(Document doc, XYZ start, XYZ end, Level level)
        {
            // 1. Define the geometry (Line)

            Line wallLine = Line.CreateBound(start, end);

            if (level == null) return null;

            try
            {
                return Wall.Create(doc, wallLine, level.Id, false);
            }
            catch (Exception ex)
            {
            }
            return null;
        }

        public Room CreateRoom(Document doc, Level level, XYZ point)
        {
            if (level == null) return null;

            try
            {
                return doc.Create.NewRoom(level, new UV(point.X, point.Y));
            }
            catch (Exception ex)
            {
            }
            return null;
        }

        private Solid CreateSolidBoudaryBuilding(Document doc, Level level, List<XYZ> lstPoint, XYZ direction, ref List<ElementId> lstDeletes, double offset = 10000 / 304.8)
        {
            Room room = CreateRoom(doc, level, (lstPoint.FirstOrDefault() + direction * offset / 5));

            if (room == null) return null;

            lstDeletes.Add(room.Id);

            List<Curve> profile = new List<Curve>();

            foreach (var listboundarySegments in room.GetBoundarySegments(new SpatialElementBoundaryOptions()))
            {
                if (listboundarySegments.Count == 4) continue;

                foreach (var segment in listboundarySegments)
                {
                    profile.Add(segment.GetCurve());
                }
            }
            CurveLoop curveLoop = CurveLoop.Create(profile);
            List<CurveLoop> loopList = new List<CurveLoop> { curveLoop };
            return GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { curveLoop }, XYZ.BasisZ, 50 / 304.8);
        }

        private List<XYZ> GetAllGeometrys(Document doc, List<Wall> wallList, out XYZ direction, out Solid solidTotal, double offset = 10000 / 304.8)
        {
            solidTotal = null;
            List<XYZ> points = new List<XYZ>();

            foreach (Wall wall in wallList)
            {
                BoundingBoxXYZ bb = wall.get_BoundingBox(doc.ActiveView);
                points.Add(bb.Max);
                points.Add(bb.Min);

                foreach (var sItem in Common.GetAllSolids(doc, wall, true))
                {
                    if (solidTotal == null)
                        solidTotal = sItem;

                    solidTotal = BooleanOperationsUtils.ExecuteBooleanOperation(solidTotal, sItem, BooleanOperationsType.Union);
                }
            }

            double xMax = points.Max(x => x.X);
            double yMax = points.Max(x => x.Y);
            double zMax = points.Max(x => x.Z);

            double xMin = points.Min(x => x.X);
            double yMin = points.Min(x => x.Y);
            double zMin = points.Min(x => x.Z);

            XYZ maxR = new XYZ(xMax, yMax, 0);

            XYZ minL = new XYZ(xMin, yMin, 0);

            direction = (maxR - minL).Normalize();

            maxR += direction * offset;
            minL -= direction * offset;

            XYZ maxL = new XYZ(minL.X, maxR.Y, 0);

            XYZ minR = new XYZ(maxR.X, minL.Y, 0);

            return new List<XYZ>() { minL, maxL, maxR, minR, minL };
        }
    }
}