using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Convert2DTo3D.Command;
using Convert2DTo3D.Utils;

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdDeleteConnectFCU : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.Doc = commandData.Application.ActiveUIDocument.Document;
            Global.AppCreation = commandData.Application.Application.Create;

            CmdConnectFCU.DeleteConnectFCUBySelection(uiDoc);

            return Result.Succeeded;
        }
    }
}