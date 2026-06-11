using Autodesk.Revit.DB;
using System.Diagnostics;

namespace Convert2DTo3D
{
    public static class DebugUtils
    {
        private static Stopwatch stopwatch = null;

        public static void DrawPoint(View view, XYZ point, Color color)
        {
            double length = 1.0;
            Line lineZ = Line.CreateBound(point + XYZ.BasisZ * length / 2.0, point - XYZ.BasisZ * length / 2.0);
            Line lineX = Line.CreateBound(point + XYZ.BasisX * length / 2.0, point - XYZ.BasisX * length / 2.0);
            Line lineY = Line.CreateBound(point + XYZ.BasisY * length / 2.0, point - XYZ.BasisY * length / 2.0);
            Document doc = view.Document;
            var m1 = doc.Create.NewModelCurve(lineZ, lineZ.GetSketchPlane(doc));
            var m2 = doc.Create.NewModelCurve(lineX, lineX.GetSketchPlane(doc));
            var m3 = doc.Create.NewModelCurve(lineY, lineY.GetSketchPlane(doc));
            SetModelEdgeColor(m1, view, color);
            SetModelEdgeColor(m2, view, color);
            SetModelEdgeColor(m3, view, color);
        }

        public static void DrawCurve(View view, Curve curve, Color color)
        {
            Document doc = view.Document;
            var m1 = doc.Create.NewModelCurve(curve, curve.GetSketchPlane(doc));
            SetModelEdgeColor(m1, view, color);
        }

        public static void DrawDetailCurve(View view, Curve curve, Color color)
        {
            Document doc = view.Document;
            var m1 = doc.Create.NewDetailCurve(view, curve);
            SetModelEdgeColor(m1, view, color);
        }

        public static void SetModelEdgeColor(Element ele, View view, Color color)
        {
            var overrides = view.GetElementOverrides(ele.Id);
            overrides.SetProjectionLineColor(color);
            view.SetElementOverrides(ele.Id, overrides);
        }

        private static void InitTime()
        {
            if (stopwatch == null) stopwatch = new Stopwatch();
        }

        public static void StartTime(string message)
        {
            InitTime();
            if (string.IsNullOrEmpty(message) == false) Debug.WriteLine(message);
            stopwatch.Restart();
            stopwatch.Start();
        }

        public static void StopTime(string message)
        {
            InitTime();
            stopwatch.Stop();
            Debug.WriteLine("{0}: {1}ms", message, stopwatch.ElapsedMilliseconds);
            stopwatch.Reset();
            stopwatch.Start();
        }

        public static void DisposeTime()
        {
            if (stopwatch != null)
            {
                stopwatch.Stop();
                stopwatch.Reset();
            }
            stopwatch = null;
        }
    }
}