using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D
{
    public static class XYZExtension
    {
        public static XYZ Round(this XYZ xYZ, int decimals = 5)
        {
            return new XYZ(Math.Round(xYZ.X, decimals), Math.Round(xYZ.Y, decimals), Math.Round(xYZ.Z, decimals));
        }

        public static string Round2Str(this XYZ xYZ, int decimals = 5)
        {
            return xYZ.Round(decimals).ToString();
        }

        public static double AngleTo2(this XYZ source, XYZ dest, XYZ upvec)
        {
            var v1 = source.Normalize();
            var v2 = dest.Normalize();
            var up = upvec.Normalize();
            var cross = v1.CrossProduct(v2);
            var dot = v1.DotProduct(v2);
            var angle = Math.Atan2(cross.GetLength(), dot);
            var test = up.DotProduct(cross);
            if (test < 0.0) angle = -angle;
            if (angle < 0) angle = 2 * Math.PI + angle;
            return angle;
        }

        private static double Area(XYZ p1, XYZ p2, XYZ p3, XYZ normal)
        {
            XYZ v = p2 - p1;
            XYZ w = p3 - p1;
            var area = v.CrossProduct(w).GetLength();
            area = area / 2.0;
            XYZ vN = v.Normalize();
            XYZ wN = w.Normalize();
            area = vN.AngleTo2(wN, normal) > Math.PI ? -area : area;
            return area;
        }

        public static bool IsLeft(this XYZ pA, XYZ pB, XYZ pCheck, XYZ normal)
        {
            double area = Area(pA, pB, pCheck, normal);
            if (Math.Abs(area - 0) < 0.00001) area = 0;
            return area > 0.0 ? true : false;
        }

        /// <summary>
        /// Chiếu các điểm point lên đường line kết quả trả về là phạm vi của điểm points trên đường line -> Line
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static Line Project2Line(this IEnumerable<XYZ> points, Line line)
        {
            if (!(points?.Count() >= 2)) return null;
            XYZ origin = line.GetEndPoint(0);
            var lineEx = line.Extend();
            SortedList<double, XYZ> mapPoints = new SortedList<double, XYZ>();
            XYZ startPoint = lineEx.GetEndPoint(0);
            double distance = 0;
            foreach (var point in points)
            {
                var newpoint = lineEx.Project(point).XYZPoint;
                distance = startPoint.DistanceTo(newpoint);
                if (mapPoints.ContainsKey(distance) == false) mapPoints.Add(distance, newpoint);
            }
            var pointSorted = mapPoints.Values.ToList();
            if (!(pointSorted?.Count() >= 2)) return null;
            startPoint = pointSorted.First();
            var endPoint = pointSorted.Last();
            if (origin.IsAlmostEqualTo(startPoint)) return Line.CreateBound(startPoint, endPoint);
            if (origin.IsAlmostEqualTo(endPoint)) return Line.CreateBound(endPoint, startPoint);
            return Line.CreateBound(startPoint, endPoint);
        }

        public static XYZ Center(this IEnumerable<XYZ> points)
        {
            if (!(points?.Count() >= 2)) return null;
            Dictionary<string, XYZ> mapPoints = new Dictionary<string, XYZ>();
            foreach (var item in points)
            {
                var s1 = item.Round2Str();
                if (mapPoints.ContainsKey(s1) == false) mapPoints.Add(s1, item);
            }
            var point2s = mapPoints.Values.ToList();
            double centerX = 0;
            double centerY = 0;
            double centerZ = 0;
            foreach (var item in point2s)
            {
                centerX += item.X;
                centerY += item.Y;
                centerZ += item.Z;
            }
            double count = point2s.Count() * 1.0;
            return new XYZ(centerX / count, centerY / count, centerZ / count);
        }
    }
}