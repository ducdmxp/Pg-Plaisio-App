using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Media3D;

namespace Convert2DTo3D.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CmdCheckFireByRadius : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            FilledRegion filledRegion = new FilteredElementCollector(doc, doc.ActiveView.Id)
              .OfClass(typeof(FilledRegion))
              .ToElements()
              .Cast<FilledRegion>().FirstOrDefault();

            if (filledRegion == null) return Result.Cancelled;

            FilledRegionType regionType = doc.GetElement(filledRegion.GetTypeId()) as FilledRegionType;

            if (regionType == null) return Result.Cancelled;

            double height = 5000 / 304.8;

            XYZ location = uiDoc.Selection.PickPoint("Pick point:");

            Transaction tran = new Transaction(doc, "Check Fire By Radius");

            try
            {
                tran.Start();

                FilledRegionType regionBlue = GetOrCreateFilledRegionType(regionType, "Koso_RegionGreen", new Color(0, 255, 0));

                FilledRegionType regionRed = GetOrCreateFilledRegionType(regionType, "Koso_RegionRed", new Color(255, 0, 0));

                Solid solidRegion = CreateSolidFromRegion(filledRegion, height);
                Solid solid1 = CreateSolid(location, 11000 / 304.8, height);
                Solid solid2 = CreateSolid(location + XYZ.BasisX * 18000 / 304.8, 9500 / 304.8, height);

                List<Solid> solids = new List<Solid>() { solid1, solid2 };

                CheckIntersecSolid(doc, regionBlue, regionRed, solidRegion, solids);

                tran.Commit();
            }
            catch (Exception)
            {
                tran.RollBack();
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private void CheckIntersecSolid(Document doc, FilledRegionType regionBlue, FilledRegionType regionRed, Solid solid, List<Solid> lstSolid)
        {
            if (solid == null || lstSolid.Count < 1) return;

            Solid totalSolidInSide = lstSolid.FirstOrDefault();

            for (int i = 0; i < lstSolid.Count; i++)
            {
                Solid solidInSide = BooleanOperationsUtils.ExecuteBooleanOperation(solid, lstSolid[i], BooleanOperationsType.Intersect);

                FilledRegion regionInSide = CreateRegion(doc, regionBlue, solidInSide);

                totalSolidInSide = BooleanOperationsUtils.ExecuteBooleanOperation(totalSolidInSide, lstSolid[i], BooleanOperationsType.Union);
            }

            Solid solidOutSide = BooleanOperationsUtils.ExecuteBooleanOperation(solid, totalSolidInSide, BooleanOperationsType.Difference);

            FilledRegion regionOutSide = CreateRegion(doc, regionRed, solidOutSide);
        }

        private FilledRegion CreateRegion(Document doc, FilledRegionType regionType, Solid solid)
        {
            if (doc == null || regionType == null) return null;

            PlanarFace planarFace = Common.GetSolidFaces(solid).Where(x => x is PlanarFace && Common.IsParallel(x.FaceNormal, XYZ.BasisZ)).OrderBy(x => x.Origin.Z).FirstOrDefault();

            if (planarFace == null) return null;

            List<CurveLoop> curveLoops = planarFace.GetEdgesAsCurveLoops().ToList();

            return FilledRegion.Create(doc, regionType.Id, doc.ActiveView.Id, curveLoops);
        }

        private Solid CreateSolid(XYZ location, double radius = 10000 / 304.8, double height = 10000)
        {
            Arc arc1 = Arc.Create(location, radius, 0, Math.PI, XYZ.BasisX, XYZ.BasisY);
            Arc arc2 = Arc.Create(location, radius, Math.PI, 2 * Math.PI, XYZ.BasisX, XYZ.BasisY);

            CurveLoop curves = new CurveLoop();
            curves.Append(arc1);
            curves.Append(arc2);

            var lstCurveLoop = new List<CurveLoop>() { curves };

            return GeometryCreationUtilities.CreateExtrusionGeometry(lstCurveLoop, XYZ.BasisZ, height);
        }

        private Solid CreateSolidFromRegion(FilledRegion filledRegion, double height = 10000)
        {
            var profileLoops = filledRegion.GetBoundaries();

            return GeometryCreationUtilities.CreateExtrusionGeometry(profileLoops, XYZ.BasisZ, height);
        }

        public FilledRegionType GetOrCreateFilledRegionType(FilledRegionType sourceTypes, string newName, Color color)
        {
            Document doc = sourceTypes.Document;

            FilledRegionType existingType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(x => x.Name.Equals(newName, System.StringComparison.InvariantCultureIgnoreCase));

            if (existingType != null) return existingType;

            FillPatternElement fillPatternElement = new FilteredElementCollector(doc)
              .OfClass(typeof(FillPatternElement))
              .Cast<FillPatternElement>()
              .FirstOrDefault(x => x.Name.ToLower().Contains("solid") || x.Name.ToLower().Contains("ソリッド"));

            FilledRegionType newType = sourceTypes.Duplicate(newName) as FilledRegionType;

            newType.BackgroundPatternColor = color;
            newType.ForegroundPatternColor = color;
            if (fillPatternElement != null)
                newType.ForegroundPatternId = fillPatternElement.Id;

            return newType;
        }

        private List<Solid> GetSolids(List<XYZ> points, double height)
        {
            List<Solid> solids = new List<Solid>();

            foreach (XYZ point in points)
            {
                Solid solid = CreateSolid(point, 7500 / 304.8, height);
                solids.Add(solid);
            }

            return solids;
        }

        private void SetElementGraphics(View activeView, ElementId elementId, Color color)
        {
            if (activeView == null || elementId == null) return;

            OverrideGraphicSettings ogs = new OverrideGraphicSettings();

            ogs.SetProjectionLineColor(color);
            ogs.SetCutBackgroundPatternColor(color);
            ogs.SetCutForegroundPatternColor(color);
            ogs.SetCutLineColor(color);

            activeView.SetElementOverrides(elementId, ogs);
        }
    }

    public class DataFire
    {
        public double Radius { get; set; }

        public double Height { get; set; }

        public XYZ Location { get; set; }

        public FamilyInstance Instance { get; set; }

        public DataFire(FamilyInstance instance)
        {
            Instance = instance;
            Height = 10000;
            Location = ((LocationPoint)Instance.Location).Point;
            Radius = 7500 / 304.8;
        }
    }
}