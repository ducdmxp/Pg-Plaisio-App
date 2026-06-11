using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Convert2DTo3D.Handler;

namespace Convert2DTo3D
{
    public class RequestHandler : IExternalEventHandler
    {
        public Request Request { get; } = new Request();

        public void Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp?.ActiveUIDocument;
            Document doc = uiDoc?.Document;

            switch (Request.Take())
            {
                case RequestId.None:
                    {
                        return;
                    }
                case RequestId.Apply:
                    {
                        return;
                    }
                default:
                    {
                        return;
					}
			}

            
        }

        public string GetName()
        {
            return "REVModelsCheck";
        }
    }
}