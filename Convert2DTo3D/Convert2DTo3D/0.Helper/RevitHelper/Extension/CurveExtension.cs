using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D
{
    public static class CurveExtension
    {
        //private static double EPSINOL = 0.0000001;

        public static List<Curve> SortContiguous(this ICollection<Curve> originalLoop)
        {
            if (originalLoop?.Count() <= 0) return null;
            List<Curve> result = new List<Curve>();
            if (originalLoop.Count == 0) return null;
            result.Add(originalLoop.First());
            IList<Curve> sorted = new List<Curve>();
            sorted.Add(originalLoop.First());

            for (int i = 0; i <= originalLoop.Count; i++)
            {
                foreach (Curve oCurve in originalLoop)
                {
                    if (sorted.Contains(oCurve)) continue;
                    if (result.Last().GetEndPoint(1).IsAlmostEqualTo(oCurve.GetEndPoint(0)))
                    {
                        sorted.Add(oCurve);
                        result.Add(oCurve);
                        break;
                    }
                    else if (result.Last().GetEndPoint(1).IsAlmostEqualTo(oCurve.GetEndPoint(1)))
                    {
                        result.Add(oCurve.CreateReversed());
                        sorted.Add(oCurve);

                        break;
                    }
                    else continue;
                }
            }
            return result.Cast<Curve>().ToList();
        }

        public static Plane GetPlane(this Curve curve)
        {
            if (!(curve is Line || curve is Arc || curve is Ellipse)) return null;
            if (curve is Line line) return Plane.CreateByNormalAndOrigin(line.GetNormal(), line.GetEndPoint(0));
            if (curve is Arc arc) return Plane.CreateByNormalAndOrigin(arc.GetNormal(), arc.Center);
            if (curve is Ellipse ellipse) return Plane.CreateByNormalAndOrigin(ellipse.GetNormal(), ellipse.Center);
            return null;
        }

        public static SketchPlane GetSketchPlane(this Curve curve, Document doc)
        {
            Plane pl = curve.GetPlane();
            if (pl == null) return null;
            return SketchPlane.Create(doc, pl);
        }

        public static XYZ GetNormal(this Line line)
        {
            var direct = line.Direction;
            XYZ normal = (!(direct.IsAlmostEqualTo(-XYZ.BasisZ) || direct.IsAlmostEqualTo(XYZ.BasisZ)))
                    ? direct.CrossProduct(XYZ.BasisZ)
                    : direct.CrossProduct(XYZ.BasisX);
            return normal;

            //XYZ ptA = line.GetEndPoint(0);
            //XYZ ptB = line.GetEndPoint(1);
            //double distance = ptA.DistanceTo(ptB);
            //if (distance <= EPSINOL) throw new Exception("Length too short.");
            //XYZ normal = ptA.CrossProduct(ptB);
            //if (normal.GetLength() < EPSINOL)
            //{
            //    XYZ aSubB = ptA.Subtract(ptB);
            //    XYZ aSubBcrossZ = aSubB.CrossProduct(XYZ.BasisZ);
            //    double crosslenth = aSubBcrossZ.GetLength();
            //    normal = (crosslenth <= EPSINOL) ? XYZ.BasisY : XYZ.BasisZ;
            //}
            // return normal;
        }

        public static XYZ GetNormal(this Arc arc)
        {
            return arc.Normal;
        }

        public static XYZ GetNormal(this Ellipse ellipse)
        {
            return ellipse.Normal;
        }

        //public static XYZ GetNormal(this CylindricalHelix cylindricalHelix)
        //{
        //    return hermiteSpline.;
        //}
        //public static XYZ GetNormal(this HermiteSpline hermiteSpline)
        //{
        //    return hermiteSpline.;
        //}
        //public static XYZ GetNormal(this NurbSpline nurbSpline)
        //{
        //    return nurbSplin;
        //}
        public static Line Extend(this Line line, double value = 10000)
        {
            return Line.CreateBound(line.GetEndPoint(0) - line.Direction * value, line.GetEndPoint(1) + line.Direction * value);
        }
    }
}