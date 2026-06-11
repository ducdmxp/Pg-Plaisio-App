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

namespace Convert2DTo3D.Commands
{
    [Transaction(TransactionMode.Manual)]
    internal class Convert2DTo3DCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var symbolId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_Walls));

            List<Line> lines = SelectLines(uiDoc);
            if (lines.Count == 0)
            {
                TaskDialog.Show("Convert 2D to 3D", "No lines selected.");
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        public static List<Line> SelectLines(UIDocument uidoc)
        {
            List<Line> curves = new List<Line>();

            try
            {
                curves = uidoc.Selection.PickObjects(ObjectType.Element, new DWGLineSelectionFilter(), "Select lines from the DWG :").Select(reference =>
                {
                    Element element = uidoc.Document.GetElement(reference);
                    if (element is CurveElement curveElement)
                    {
                        return curveElement.GeometryCurve as Line;
                    }
                    return null;
                }).Where(line => line != null).ToList();
            }
            catch (Exception)
            {
                throw;
            }

            return curves;
        }
    }
}