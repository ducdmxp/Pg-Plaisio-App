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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var symbolId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_Walls));

            List<ModelLine> lines = SelectLines(uiDoc);
            if (lines.Count == 0)
            {
                TaskDialog.Show("Convert 2D to 3D", "No lines selected.");
                return Result.Cancelled;
            }

            List<ModelLine> tuongchinhs = GetLinesByStyle(lines, "NOSIVO");

            uiDoc.Selection.SetElementIds(tuongchinhs.Select(line => line.Id).ToList());

            return Result.Succeeded;
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