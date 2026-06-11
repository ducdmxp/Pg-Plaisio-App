#region Name spaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Convert2DTo3D.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using ParameterUtils = Convert2DTo3D.Utils.ParameterUtils;

#endregion Name spaces

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdConnectWC : IExternalCommand
    {
        public const string MEP_Storage_ConnectWC = "MEP_Storage_ConnectWC";

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

            ConnectWCFrm form = new ConnectWCFrm();
            if (form.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            List<RevitLinkInstance> ListRevitLinkInstances = new FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_RvtLinks)
                 .OfClass(typeof(RevitLinkInstance))
                 .Cast<RevitLinkInstance>()
                 .Where(x => x.GetLinkDocument() != null)
                 .OrderBy(x => x.Name).ToList();

            Element elePicked = PickElement(uiDoc, form.IsPickWall, form.IsProject);

            if (elePicked == null)
            {
                return Result.Cancelled;
            }

            TransactionGroup tranG = new TransactionGroup(doc, "ConnectPipes");

            try
            {
                tranG.Start();

                int SelectedType = form.SelectedType;

                double offset = form.Offset / 304.8;
                double offsetFromLevel = form.OffsetLevel / 304.8;
                double diamter = form.Diameter / 304.8;
                bool isElbow45 = form.IsElbow45;

                FamilySymbol symbolCoren = form.SymbolCoren;

                if (SelectedType == 0)
                {
                    Plane planePick = GetPlane(elePicked, ListRevitLinkInstances, out Line lcLine);

                    List<Pipe> lstPipePickeds = PickPipes(uiDoc, out Line lineMEPPicked);

                    List<XYZ> lstPointPicked = SortPointsByDirection(PickPoint(uiDoc), lineMEPPicked.Direction, lineMEPPicked.GetEndPoint(0));

                    lstPointPicked = lstPointPicked.Select(x => Common.GetPointProjectOnPlane(planePick, x, null)).ToList();

                    foreach (var item in lstPipePickeds)
                    {
                        Pipe mepCurvePicked = item;

                        Line lineMEP = (mepCurvePicked.Location as LocationCurve).Curve as Line;

                        for (int i = 0; i < lstPointPicked.Count; i++)
                        {
                            XYZ point = lstPointPicked[i];

                            if (form.IsTee)
                            {
                                if (IsPointOnPipe(mepCurvePicked, point) == false)
                                    continue;
                            }
                            else
                            {
                                if (i == 0 || i == lstPointPicked.Count - 1)
                                {
                                }
                                else
                                {
                                    if (IsPointOnPipe(mepCurvePicked, point) == false)
                                        continue;
                                }
                            }

                            if ((i == 0 || i == lstPointPicked.Count - 1) && form.IsTee == false
                                && !IsConnectedEnd(mepCurvePicked, out Connector conNotConnec)
                                && IsCreateElbow(lineMEP, point, conNotConnec))
                            {
                                CreateSystemWCType2(doc, ref mepCurvePicked, point, elePicked, planePick,
                                                    lcLine, symbolCoren, offset, offsetFromLevel, isElbow45);
                            }
                            else if (IsPointOnPipe(mepCurvePicked, point))
                            {
                                CreateSystemWCType1(doc, ref mepCurvePicked, point, elePicked, planePick,
                                                    lcLine, symbolCoren, lineMEPPicked.GetEndPoint(0),
                                                    offset, offsetFromLevel, diamter, isElbow45);
                            }
                        }
                    }
                }
                else if (SelectedType == 1)
                {
                    //Element elePicked = PickElement(uiDoc, form.IsPickWall, form.IsProject);

                    List<Pipe> lstPipePickeds = PickPipes(uiDoc, out Line lineMEPPicked);

                    foreach (var item in lstPipePickeds)
                    {
                        Pipe mepCurvePicked = item;

                        Plane planePick = GetPlane(elePicked, ListRevitLinkInstances, out Line lcLine);

                        Line lineMEP = (mepCurvePicked.Location as LocationCurve).Curve as Line;

                        Connector con = ConnectorUtils.GetConnectorNotConnnected(mepCurvePicked.ConnectorManager);

                        if (con != null)
                        {
                            XYZ point = con.Origin + con.CoordinateSystem.BasisZ * 100 / 304.8;

                            CreateSystemWCType31(doc, ref mepCurvePicked, point, elePicked, planePick,
                                     lcLine, symbolCoren, offset, offsetFromLevel, diamter, isElbow45);
                        }
                    }
                }
                else
                {
                    //Element elePicked = PickElement(uiDoc, form.IsPickWall, form.IsProject);

                    Plane planePick = GetPlane(elePicked, ListRevitLinkInstances, out Line lcLine);

                    Pipe mepCurvePicked = doc.GetElement(uiDoc.Selection.PickObject(ObjectType.Element, new PipeSelectionFilter(), "Select a pipe")) as Pipe;

                    Line lineMEP = (mepCurvePicked.Location as LocationCurve).Curve as Line;

                    XYZ point = uiDoc.Selection.PickPoint("Pick a point on the pipe");

                    CreateSystemWCType32(doc, ref mepCurvePicked, point, elePicked, planePick,
                        lcLine, symbolCoren, offset, offsetFromLevel, diamter, isElbow45);
                }

                tranG.Assimilate();
            }
            catch (Exception ex)
            {
                if (tranG.HasStarted())
                    tranG.RollBack();
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private void CreateSystemWCType1(Document doc, ref Pipe pipePick, XYZ pointPicked, Element elePicked,
            Plane planePick, Line lcLine, FamilySymbol symbolCoren, XYZ origin,
            double offset = 200 / 304.8, double offsetFromLevel = 250 / 304.8,
            double diamter = 20 / 304.8,
            bool isElbow45 = true)
        {
            Transaction tran = new Transaction(doc, "ConnectPipes");

            try
            {
                ElementId systemTypeId = ParameterUtils.GetValueParameterByBuilt(pipePick, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) as ElementId;

                tran.Start();

                FailureHandlingOptions options = tran.GetFailureHandlingOptions();
                DisableWarning preproccessor = new DisableWarning();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                tran.SetFailureHandlingOptions(options);

                if (!symbolCoren.IsActive)
                    symbolCoren.Activate();

                List<Element> elements = new List<Element>();

                Line lineMEP = (pipePick.Location as LocationCurve).Curve as Line;

                Plane planeMEP = Plane.CreateByNormalAndOrigin(lineMEP.Direction.CrossProduct(XYZ.BasisZ), lineMEP.GetEndPoint(0));

                XYZ pProject = Common.GetPointProjectOnPlane(planeMEP, pointPicked, null);

                XYZ vectorX = (pointPicked - Common.GetPointProjectOnPlane(planeMEP, pointPicked, null)).Normalize();

                XYZ pOnMep = Common.GetPointProjectOnLine(lineMEP, pointPicked);

                XYZ p0 = (isElbow45) ? pOnMep + vectorX * Math.Abs(offset) + XYZ.BasisZ * offset : pOnMep + XYZ.BasisZ * offset;
                XYZ p1 = Common.GetPointProjectOnPlane(planePick, p0, null);

                Pipe pipeBranch1 = null;
                if (!Common.IsEqual(offset, 0))
                {
                    pipeBranch1 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, pOnMep, p0);
                    Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch1, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                    var tee = CreateWYE(doc, ref pipePick, pipeBranch1, origin);

                    if (tee == null)
                    {
                        if (tran.HasStarted())
                            tran.RollBack();
                        return;
                    }

                    tee.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, "Tee");
                    elements.Add(pipeBranch1);
                }

                Pipe pipeBranch2 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p0, p1);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeBranch2);

                if (pipeBranch1 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch1.ConnectorManager, pipeBranch2.ConnectorManager, out Connector conBranch1, out Connector conBranch2);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(conBranch1, conBranch2);

                    elements.Add(fitting);
                }
                else
                {
                    var tee = CreateWYE(doc, ref pipePick, pipeBranch2, origin);

                    if (tee == null)
                    {
                        if (tran.HasStarted())
                            tran.RollBack();
                        return;
                    }

                    tee.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, "Tee");
                }

                XYZ p2 = p1.To2D(pipePick.ReferenceLevel.Elevation + offsetFromLevel);

                Pipe pipeVertical = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p1, p2);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeVertical, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeVertical);

                if (pipeVertical != null && pipeBranch2 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                FamilyInstance coren = CreateCoren(doc, symbolCoren, pipePick, pipeVertical, elePicked, planePick, p2, vectorX, offsetFromLevel);

                elements.Add(coren);

                foreach (var element in elements)
                {
                    if (element != null && element.IsValidObject)
                        element.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, true.ToString());
                }

                tran.Commit();
            }
            catch (Exception ex)
            {
                string message = ex.Message;

                if (tran.HasStarted())
                    tran.RollBack();
            }
        }

        private void CreateSystemWCType2(Document doc, ref Pipe pipePick, XYZ pointPick, Element elePicked,
            Plane planePick, Line lcLine, FamilySymbol symbolCoren,
            double offset = 200 / 304.8, double offsetFromLevel = 250 / 304.8,
            bool isElbow45 = true)
        {
            Transaction tran = new Transaction(doc, "ConnectPipes");

            try
            {
                ElementId systemTypeId = ParameterUtils.GetValueParameterByBuilt(pipePick, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) as ElementId;

                double diamter = pipePick.Diameter;

                tran.Start();

                FailureHandlingOptions options = tran.GetFailureHandlingOptions();
                DisableWarning preproccessor = new DisableWarning();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                tran.SetFailureHandlingOptions(options);

                if (!symbolCoren.IsActive)
                    symbolCoren.Activate();

                List<Element> elements = new List<Element>();

                Line lineMEP = (pipePick.Location as LocationCurve).Curve as Line;

                Plane planeMEP = Plane.CreateByNormalAndOrigin(lineMEP.Direction.CrossProduct(XYZ.BasisZ), lineMEP.GetEndPoint(0));

                XYZ pProject = Common.GetPointProjectOnPlane(planeMEP, pointPick, null);

                XYZ vectorX = (pointPick - Common.GetPointProjectOnPlane(planeMEP, pointPick, null)).Normalize();

                XYZ pOnMep = Common.GetPointProjectOnLine(lineMEP, pointPick);

                XYZ p0 = (isElbow45) ? pOnMep + vectorX * Math.Abs(offset) + XYZ.BasisZ * offset : pOnMep + XYZ.BasisZ * offset;
                XYZ p1 = Common.GetPointProjectOnPlane(planePick, p0, null);

                Pipe pipeBranch1 = null;
                if (!Common.IsEqual(offset, 0))
                {
                    pipeBranch1 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, pOnMep, p0);
                    Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch1, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                    Connector con01 = ConnectorUtils.GetConnectorNotConnnected(pipePick.ConnectorManager);// ConnectorUtils.GetConnectorNearest(origin, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch1.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                    elements.Add(pipeBranch1);
                }

                Pipe pipeBranch2 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p0, p1);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeBranch2);

                if (pipeBranch1 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch1.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }
                else
                {
                    Connector con01 = ConnectorUtils.GetConnectorNotConnnected(pipePick.ConnectorManager);//ConnectorUtils.GetConnectorNearest(origin, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch2.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                XYZ p2 = p1.To2D(pipePick.ReferenceLevel.Elevation + offsetFromLevel);

                Pipe pipeVertical = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p1, p2);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeVertical, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeVertical);

                if (pipeVertical != null && pipeBranch2 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                FamilyInstance coren = CreateCoren(doc, symbolCoren, pipePick, pipeVertical, elePicked, planePick, p2, vectorX, offsetFromLevel);

                elements.Add(coren);

                foreach (var element in elements)
                {
                    if (element != null && element.IsValidObject)
                        element.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, true.ToString());
                }

                tran.Commit();
            }
            catch (Exception ex)
            {
                string message = ex.Message;

                if (tran.HasStarted())
                    tran.RollBack();
            }
        }

        private void CreateSystemWCType31(Document doc, ref Pipe pipePick, XYZ pointPick, Element elePicked,
          Plane planePick, Line lcLine, FamilySymbol symbolCoren,
          double offset = 200 / 304.8, double offsetFromLevel = 250 / 304.8,
          double diamter = 20 / 304.8,
          bool isElbow45 = true)
        {
            Transaction tran = new Transaction(doc, "ConnectPipes");

            try
            {
                if (!Common.IsEqual(offset, 0))
                    diamter = pipePick.Diameter;

                ElementId systemTypeId = ParameterUtils.GetValueParameterByBuilt(pipePick, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) as ElementId;

                tran.Start();

                FailureHandlingOptions options = tran.GetFailureHandlingOptions();
                DisableWarning preproccessor = new DisableWarning();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                tran.SetFailureHandlingOptions(options);

                if (!symbolCoren.IsActive)
                    symbolCoren.Activate();

                List<Element> elements = new List<Element>();

                Line lineMEP = (pipePick.Location as LocationCurve).Curve as Line;

                Connector conMEP = ConnectorUtils.GetConnectorNearest(pointPick, pipePick.ConnectorManager, out _);

                XYZ vectorX = conMEP.CoordinateSystem.BasisZ;

                XYZ pOnMep = conMEP.Origin;

                XYZ p0 = (isElbow45) ? pOnMep + vectorX * Math.Abs(offset) + XYZ.BasisZ * offset : pOnMep + XYZ.BasisZ * offset;
                XYZ p1 = Common.GetPointProjectOnPlane(planePick, p0, null);

                Pipe pipeBranch1 = null;
                if (!Common.IsEqual(offset, 0))
                {
                    pipeBranch1 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, pOnMep, p0);
                    Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch1, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                    Connector con01 = ConnectorUtils.GetConnectorNearest(pOnMep, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch1.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                    elements.Add(pipeBranch1);
                }

                Pipe pipeBranch2 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p0, p1);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeBranch2);

                if (pipeBranch1 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch1.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }
                else
                {
                    Connector con01 = ConnectorUtils.GetConnectorNearest(pOnMep, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch2.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewTransitionFitting(con01, con02);

                    if (fitting != null)
                        elements.Add(fitting);
                    else
                    {
                        pipePick = CreateNewMepCurve(doc, pipePick, pipeBranch2) as Pipe;

                        pipeBranch2 = pipePick;
                    }
                }

                XYZ p2 = p1.To2D(pipePick.ReferenceLevel.Elevation + offsetFromLevel);

                Pipe pipeVertical = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p1, p2);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeVertical, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeVertical);

                if (pipeVertical != null && pipeBranch2 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                FamilyInstance coren = CreateCoren(doc, symbolCoren, pipePick, pipeVertical, elePicked, planePick, p2, vectorX, offsetFromLevel);
                elements.Add(coren);

                foreach (var element in elements)
                {
                    if (element != null && element.IsValidObject)
                        element.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, true.ToString());
                }

                tran.Commit();
            }
            catch (Exception ex)
            {
                string message = ex.Message;

                if (tran.HasStarted())
                    tran.RollBack();
            }
        }

        private void CreateSystemWCType32(Document doc, ref Pipe pipePick, XYZ pointPick, Element elePicked,
         Plane planePick, Line lcLine, FamilySymbol symbolCoren,
         double offset = 200 / 304.8, double offsetFromLevel = 250 / 304.8,
         double diamter = 20 / 304.8,
         bool isElbow45 = true, double lengthDefault = 500 / 304.8)
        {
            Transaction tran = new Transaction(doc, "ConnectPipes");

            try
            {
                if (!Common.IsEqual(offset, 0))
                    diamter = pipePick.Diameter;

                ElementId systemTypeId = ParameterUtils.GetValueParameterByBuilt(pipePick, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) as ElementId;

                tran.Start();

                FailureHandlingOptions options = tran.GetFailureHandlingOptions();
                DisableWarning preproccessor = new DisableWarning();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                tran.SetFailureHandlingOptions(options);

                if (!symbolCoren.IsActive)
                    symbolCoren.Activate();

                List<Element> elements = new List<Element>();

                Line lineMEP = (pipePick.Location as LocationCurve).Curve as Line;

                Connector conMEP = ConnectorUtils.GetConnectorNearest(pointPick, pipePick.ConnectorManager, out _);

                XYZ vectorX = conMEP.CoordinateSystem.BasisZ;

                XYZ pOnMep = conMEP.Origin;

                XYZ vectorY = (pointPick - Common.GetPointProjectOnLine(lineMEP, pointPick)).To2D();

                Plane planePointPick = Plane.CreateByNormalAndOrigin(vectorY, pointPick);

                XYZ p0 = (isElbow45) ? pOnMep + vectorX * Math.Abs(offset) + XYZ.BasisZ * offset : pOnMep + XYZ.BasisZ * offset;
                XYZ p1 = p0 + vectorX * lengthDefault;// Common.GetPointProjectOnPlane(planePick, p0, null);

                Pipe pipeBranch1 = null;
                if (!Common.IsEqual(offset, 0))
                {
                    pipeBranch1 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, pOnMep, p0);
                    Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch1, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                    Connector con01 = ConnectorUtils.GetConnectorNearest(pOnMep, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch1.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                    elements.Add(pipeBranch1);
                }

                Pipe pipeBranch2 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p0, p1);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);
                elements.Add(pipeBranch2);

                if (pipeBranch1 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch1.ConnectorManager, pipeBranch2.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);
                    elements.Add(fitting);
                }
                else
                {
                    Connector con01 = ConnectorUtils.GetConnectorNearest(pOnMep, pipePick.ConnectorManager, out _);

                    Connector con02 = ConnectorUtils.GetConnectorNearest(pOnMep, pipeBranch2.ConnectorManager, out _);

                    FamilyInstance fitting = doc.Create.NewTransitionFitting(con01, con02);

                    if (fitting != null)
                        elements.Add(fitting);
                    else
                    {
                        pipePick = CreateNewMepCurve(doc, pipePick, pipeBranch2) as Pipe;

                        pipeBranch2 = pipePick;
                    }
                }

                XYZ p2 = Common.GetPointProjectOnPlane(planePointPick, p1, null);

                Pipe pipeBranch3 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p1, p2);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch3, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeBranch3);

                if (pipeBranch2 != null && pipeBranch3 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch2.ConnectorManager, pipeBranch3.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                XYZ p3 = Common.GetPointProjectOnPlane(planePick, p2, null);

                Pipe pipeBranch4 = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p2, p3);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch4, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeBranch4);

                if (pipeBranch3 != null && pipeBranch4 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeBranch3.ConnectorManager, pipeBranch4.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);

                    elements.Add(fitting);
                }

                XYZ p4 = p3.To2D(pipePick.ReferenceLevel.Elevation + offsetFromLevel);

                Pipe pipeVertical = Pipe.Create(doc, systemTypeId, pipePick.GetTypeId(), pipePick.ReferenceLevel.Id, p3, p4);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeVertical, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, diamter);

                elements.Add(pipeVertical);

                if (pipeVertical != null && pipeBranch4 != null)
                {
                    ConnectorUtils.GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeBranch4.ConnectorManager, out Connector con01, out Connector con02);

                    FamilyInstance fitting = doc.Create.NewElbowFitting(con01, con02);
                    elements.Add(fitting);
                }

                FamilyInstance coren = CreateCoren(doc, symbolCoren, pipePick, pipeVertical, elePicked, planePick, p4, vectorX, offsetFromLevel);

                elements.Add(coren);

                foreach (var element in elements)
                {
                    if (element != null && element.IsValidObject)
                        element.SetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC, true.ToString());
                }

                tran.Commit();
            }
            catch (Exception ex)
            {
                string message = ex.Message;

                if (tran.HasStarted())
                    tran.RollBack();
            }
        }

        private FamilyInstance CreateCoren(Document doc, FamilySymbol symbolCoren, Pipe pipePick,
            Pipe pipeBranch3, Element elePicked, Plane planePick, XYZ p2, XYZ vectorX, double elevationFromLevel)
        {
            FamilyInstance fitting = doc.Create.NewFamilyInstance(p2, symbolCoren, pipePick.ReferenceLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            doc.Regenerate();

            List<Connector> connectors = ConnectorUtils.ToList(fitting.MEPModel.ConnectorManager).OrderBy(x => x.Radius).ToList();

            Connector conCoren1 = connectors.FirstOrDefault();

            Connector con1 = ConnectorUtils.GetConnectorNotConnnected(pipeBranch3.ConnectorManager);

            Connector conCoren2 = connectors.LastOrDefault();

            double angle = conCoren1.CoordinateSystem.BasisZ.AngleTo(con1.CoordinateSystem.BasisZ);

            Line lineAxis = Line.CreateUnbound(conCoren2.Origin, conCoren2.CoordinateSystem.BasisZ);

            ElementTransformUtils.RotateElement(doc, fitting.Id, lineAxis, angle);

            ElementTransformUtils.MoveElement(doc, fitting.Id, con1.Origin - conCoren1.Origin);

            if (conCoren1.CoordinateSystem.BasisZ.DotProduct(con1.CoordinateSystem.BasisZ) > 0)
                ElementTransformUtils.RotateElement(doc, fitting.Id, lineAxis, Math.PI);

            conCoren1.ConnectTo(con1);

            angle = conCoren2.CoordinateSystem.BasisZ.AngleTo(vectorX.Negate());

            ElementTransformUtils.RotateElement(doc, fitting.Id, Line.CreateUnbound(con1.Origin, con1.CoordinateSystem.BasisZ), angle);

            doc.Regenerate();

            if (vectorX.DotProduct(conCoren2.CoordinateSystem.BasisZ) > 0)
                ElementTransformUtils.RotateElement(doc, fitting.Id, Line.CreateUnbound(con1.Origin, con1.CoordinateSystem.BasisZ), Math.PI);

            doc.Regenerate();

            Plane planeCoren = GetFaceFitting(fitting);

            double widthOffset = (elePicked is Wall wall) ? wall.Width / 2 : 0;

            Plane planar = Plane.CreateByNormalAndOrigin(planePick.Normal, planePick.Origin + vectorX.Negate() * widthOffset);

            ElementTransformUtils.MoveElement(doc, fitting.Id, Common.GetPointProjectOnPlane(planar, planeCoren.Origin, null) - planeCoren.Origin);

            doc.Regenerate();
            ParameterUtils.SetValueParameterByBuiltIn(fitting, BuiltInParameter.INSTANCE_ELEVATION_PARAM, elevationFromLevel);

            return fitting;
        }

        private List<XYZ> PickPoint(UIDocument uidoc)
        {
            List<XYZ> points = new List<XYZ>();
            try
            {
                while (true)
                {
                    try
                    {
                        XYZ point = uidoc.Selection.PickPoint("Pick a point :");
                        if (point != null)
                            points.Add(point);
                        else
                            break;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled the selection
            }
            return points;
        }

        private Element PickElement(UIDocument uidoc, bool isPickWall = true, bool isProject = true)
        {
            try
            {
                if (isPickWall)
                {
                    List<Category> categories = new List<Category>() { Category.GetCategory(uidoc.Document, BuiltInCategory.OST_Walls) };
                    if (isProject)
                    {
                        return uidoc.Document.GetElement(uidoc.Selection.PickObject(ObjectType.Element, new TypeSelectionFilter(categories), "Select a wall :"));
                    }
                    else
                    {
                        Reference item = uidoc.Selection.PickObject(ObjectType.LinkedElement, new TypeSelectionFilter(categories), "Select a wall :");

                        RevitLinkInstance link = uidoc.Document.GetElement(item) as RevitLinkInstance;

                        return link.GetLinkDocument().GetElement(item.LinkedElementId) as Element;
                    }
                }
                else
                {
                    List<Category> categories = new List<Category>() { Category.GetCategory(uidoc.Document, BuiltInCategory.OST_Lines) };
                    return uidoc.Document.GetElement(uidoc.Selection.PickObject(ObjectType.Element, new TypeSelectionFilter(categories), "Select a ModelLine or DetailLine :"));
                }
            }
            catch (Exception)
            {
            }
            return null;
        }

        private Plane GetPlane(Element ele, List<RevitLinkInstance> revitLinkInstances, out Line lcLine)
        {
            lcLine = null;

            if (ele != null)
            {
                Transform transform = null;
                if (revitLinkInstances.Count > 0 && ele.Document.IsLinked)
                {
                    RevitLinkInstance revitLinkInstance = revitLinkInstances.FirstOrDefault(x => x.GetLinkDocument().Title == ele.Document.Title);
                    if (revitLinkInstance != null)
                        transform = revitLinkInstance.GetTotalTransform();
                }

                if (ele.Location is LocationCurve lc)
                {
                    lcLine = ((LocationCurve)ele.Location).Curve as Line;

                    XYZ p0 = lcLine.GetEndPoint(0);
                    XYZ p1 = lcLine.GetEndPoint(1);

                    if (transform != null)
                    {
                        p0 = transform.OfPoint(p0);
                        p1 = transform.OfPoint(p1);
                    }

                    lcLine = Line.CreateBound(p0, p1);

                    XYZ normal = Common.To2D(lcLine.Direction).CrossProduct(XYZ.BasisZ).Normalize();

                    return Plane.CreateByNormalAndOrigin(normal, p0);
                }
            }

            return null;
        }

        private Plane GetFaceFitting(FamilyInstance instance)
        {
            if (instance == null || instance.Document == null)
                return null;

            var solids = Common.GetAllSolids(instance.Document, instance, true).ToList();

            List<Arc> arcs = new List<Arc>();
            foreach (var solid in solids)
            {
                arcs.AddRange(Common.GetAllCurves(solid).Where(x => x is Arc).Cast<Arc>().ToList());
            }

            Connector connector = ConnectorUtils.ToList(instance.MEPModel.ConnectorManager).FirstOrDefault(x => !Common.IsParallel(x.CoordinateSystem.BasisZ, XYZ.BasisZ));

            List<Arc> lstArcSorts = SortArcByDirection(arcs, connector.CoordinateSystem.BasisZ, connector.Origin);

            Arc arc = lstArcSorts.LastOrDefault();

            return Plane.CreateByNormalAndOrigin(arc.Normal, arc.Center);
        }

        public List<Arc> SortArcByDirection(List<Arc> lstArcs, XYZ direction, XYZ referencePoint = null)
        {
            if (lstArcs == null || !lstArcs.Any())
                return new List<Arc>();

            if (direction == null || direction.IsZeroLength())
                throw new ArgumentException("Vector hướng không hợp lệ");

            if (referencePoint == null)
                referencePoint = XYZ.Zero;

            XYZ normalizedDirection = direction.Normalize();

            lstArcs = lstArcs.Where(x => (x.Center - referencePoint).DotProduct(normalizedDirection) > 0).ToList();

            var sortedPoints = lstArcs.OrderBy(x =>
            {
                XYZ vectorToPoint = x.Center - referencePoint;

                double projection = vectorToPoint.DotProduct(normalizedDirection);

                return projection;
            }).ThenBy(x => x.Radius).ToList();

            return sortedPoints;
        }

        public List<XYZ> SortPointsByDirection(List<XYZ> points, XYZ direction, XYZ referencePoint = null)
        {
            if (points == null || !points.Any())
                return new List<XYZ>();

            if (direction == null || direction.IsZeroLength())
                throw new ArgumentException("Vector hướng không hợp lệ");

            if (referencePoint == null)
                referencePoint = XYZ.Zero;

            XYZ normalizedDirection = direction.Normalize();

            var sortedPoints = points.OrderBy(point =>
            {
                XYZ vectorToPoint = point - referencePoint;

                double projection = vectorToPoint.DotProduct(normalizedDirection);

                return projection;
            }).ToList();

            return sortedPoints;
        }

        public FamilyInstance CreateWYE(Document doc, ref Pipe mEPCurveMain1, Pipe mEPCurveBranch, XYZ origin)
        {
            FamilyInstance fittingWye = null;

            try
            {
                ElementId systemTypeId = ParameterUtils.GetValueParameterByBuilt(mEPCurveMain1, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) as ElementId;

                Line lineMain1 = (mEPCurveMain1.Location as LocationCurve).Curve as Line;

                Line lineBranch = (mEPCurveBranch.Location as LocationCurve).Curve as Line;

                XYZ p = Common.GetPointIntersecNotInXYPlane(lineMain1, lineBranch, true);

                ElementId elementId = BreakMEPCurveUtils.BreakMEPCurve(doc, mEPCurveMain1.Id, p);

                Pipe mEPCurveMain2 = doc.GetElement(elementId) as Pipe;

                if (mEPCurveMain2 == null)
                    return fittingWye;

                Line lineMain2 = (mEPCurveMain2.Location as LocationCurve).Curve as Line;

                Pipe pipeMain1 = Pipe.CreatePlaceholder(doc, systemTypeId, mEPCurveMain1.GetTypeId(),
                    mEPCurveMain1.ReferenceLevel.Id, lineMain1.GetEndPoint(0), lineMain1.GetEndPoint(1));

                Pipe pipeMain2 = Pipe.CreatePlaceholder(doc, systemTypeId, mEPCurveMain1.GetTypeId(),
                   mEPCurveMain1.ReferenceLevel.Id, lineMain2.GetEndPoint(0), lineMain2.GetEndPoint(1));

                Pipe pipeBranch = Pipe.CreatePlaceholder(doc, systemTypeId, mEPCurveBranch.GetTypeId(),
                 mEPCurveBranch.ReferenceLevel.Id, lineBranch.GetEndPoint(0), lineBranch.GetEndPoint(1));

                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeMain1, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, mEPCurveMain1.Diameter);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeMain2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, mEPCurveMain1.Diameter);
                Utils.ParameterUtils.SetValueParameterByBuiltIn(pipeBranch, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, mEPCurveBranch.Diameter);

                doc.Regenerate();

                ConnectorUtils.GetConnectorClosedTo(pipeMain1.ConnectorManager, pipeMain2.ConnectorManager, out Connector con1, out Connector con2);

                Connector con3 = ConnectorUtils.GetConnectorNearest(p, pipeBranch.ConnectorManager, out _);

                PlumbingUtils.ConnectPipePlaceholdersAtTee(doc, con1, con2, con3);

                List<ElementId> placeholders = new List<ElementId>();

                if (pipeMain1 != null && pipeMain1.IsValidObject)
                    placeholders.Add(pipeMain1.Id);

                if (pipeMain2 != null && pipeMain2.IsValidObject)
                    placeholders.Add(pipeMain2.Id);

                if (pipeBranch != null && pipeBranch.IsValidObject)
                    placeholders.Add(pipeBranch.Id);

                placeholders.Add(pipeBranch.Id);

                List<ElementId> elementIds = PlumbingUtils.ConvertPipePlaceholders(doc, placeholders).ToList();

                fittingWye = elementIds.Select(x => doc.GetElement(x)).Where(x => x is FamilyInstance).Cast<FamilyInstance>().FirstOrDefault();

                if (fittingWye == null)
                    return null;

                doc.Delete(elementIds.Where(x => doc.GetElement(x) is MEPCurve).ToList());

                Common.GetInformationConectorWye(fittingWye, null, out Connector main1, out Connector main2, out Connector conTee);

                Connector c3 = ConnectorUtils.GetConnectorNearest(p, mEPCurveBranch.ConnectorManager, out _);

                if (c3 != null && !c3.IsConnectedTo(conTee))
                    c3.ConnectTo(conTee);

                ConnectorUtils.GetConnectorClosedTo(mEPCurveMain1.ConnectorManager, mEPCurveMain2.ConnectorManager, out Connector c1, out Connector c2);

                Connector conValid1 = (c1.CoordinateSystem.BasisZ.DotProduct(main1.CoordinateSystem.BasisZ) < 0) ? main1 : main2;

                Connector conValid2 = (c2.CoordinateSystem.BasisZ.DotProduct(main1.CoordinateSystem.BasisZ) < 0) ? main1 : main2;

                if (c1 != null && !c1.IsConnectedTo(conValid1))
                    c1.ConnectTo(conValid1);

                if (c2 != null && !c2.IsConnectedTo(conValid2))
                    c2.ConnectTo(conValid2);

                XYZ vectorMove = lineMain1.Direction * 1 / 304.8;

                ElementTransformUtils.MoveElement(doc, fittingWye.Id, vectorMove);

                doc.Regenerate();

                ElementTransformUtils.MoveElement(doc, fittingWye.Id, vectorMove.Negate());

                doc.Regenerate();

                mEPCurveMain1 = (mEPCurveMain2 == null) ? mEPCurveMain1 : Common.GetNextPipe(mEPCurveMain1, mEPCurveMain2, origin);
            }
            catch (Exception ex)
            {
                return null;
            }

            return fittingWye;
        }

        public static void DeleteConnectWCBySelection(UIDocument uiDoc)
        {
            List<Element> lstElementIdSelecteds = uiDoc.Selection.GetElementIds().Select(x => uiDoc.Document.GetElement(x))
                .Where(x => IsFittingGroup(x, out _))
                .ToList();

            if (lstElementIdSelecteds == null || lstElementIdSelecteds.Count <= 0)
            {
                // IO.ShowWarning("Please select MEPCurve before running the command");
                return;
            }

            Transaction tran = new Transaction(uiDoc.Document, "DeleteMEPCurve");

            try
            {
                tran.Start();

                FailureHandlingOptions options = tran.GetFailureHandlingOptions();
                DisableWarning preproccessor = new DisableWarning();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                tran.SetFailureHandlingOptions(options);

                List<Element> elements = new List<Element>();

                foreach (var element in lstElementIdSelecteds)
                {
                    List<Element> lstElementInGroup = GetAllElementConnected(uiDoc.Document, element.Id, true);

                    elements.AddRange(lstElementInGroup);
                }

                elements = elements.GroupBy(x => x.Id).Select(x => x.FirstOrDefault()).Where(x => IsFittingGroup(x, out string value)).ToList();

                List<Element> lstMEPDeletes = elements.Where(x => IsFittingGroup(x, out string value) && !value.Equals("Tee")).ToList();

                List<FamilyInstance> lstFittings = elements.Where(x => IsFittingGroup(x, out string value) && value.Equals("Tee")).Cast<FamilyInstance>().ToList();

                uiDoc.Document.Delete(lstMEPDeletes.Select(x => x.Id).ToList());
                uiDoc.Document.Regenerate();

                foreach (var fitting in lstFittings)
                {
                    if (fitting != null && fitting.IsValidObject)
                    {
                        Common.GetInformationConectorWye(fitting as FamilyInstance, null, out Connector main1, out Connector main2, out Connector conY);

                        ConnectorUtils.DisconnectFrom(main1, out Element mEPCurve1);
                        ConnectorUtils.DisconnectFrom(main2, out Element mEPCurve2);

                        uiDoc.Document.Delete(fitting.Id);

                        uiDoc.Document.Regenerate();

                        CreateNewMepCurve(uiDoc.Document, mEPCurve1 as MEPCurve, mEPCurve2 as MEPCurve);
                    }
                }

                tran.Commit();
            }
            catch (Exception)
            {
                if (tran.HasStarted())
                    tran.RollBack();
            }
        }

        private static bool IsFittingGroup(Element element, out string value)
        {
            value = string.Empty;
            if (element != null)
            {
                string valueEntity = element.GetElementParameterDataStorage<string>(MEP_Storage_ConnectWC, MEP_Storage_ConnectWC);

                if (valueEntity != null && !string.IsNullOrEmpty(valueEntity as string))
                {
                    value = valueEntity as string;
                    return true;
                }
            }

            return false;
        }

        public static List<Element> GetAllElementConnected(Document document, ElementId elementMainId, bool addFirst = false, List<Element> rmvElement = null)
        {
            List<Element> components = new List<Element>();
            Element eleMain = document.GetElement(elementMainId);
            List<Element> rmvElement_1 = null;
            if (addFirst == true)
            {
                rmvElement_1 = new List<Element> { eleMain };
                components.Add(eleMain);
            }
            else
            {
                if (rmvElement == null)
                    rmvElement = new List<Element>();

                rmvElement_1 = new List<Element>(rmvElement);
            }
            if (eleMain != null)
            {
                List<Connector> connectors = new List<Connector>();
                if (eleMain is MEPCurve processMEPCurve && processMEPCurve.ConnectorManager != null)
                {
                    bool isConnectWC = IsFittingGroup(eleMain, out _);

                    if (isConnectWC)
                    {
                        foreach (Connector connector in processMEPCurve.ConnectorManager.Connectors)
                        {
                            //if (connector.ConnectorType != ConnectorType.End)
                            //    continue;
                            connectors.Add(connector);
                        }
                    }
                }
                else if (eleMain is FamilyInstance fmlIns && fmlIns.MEPModel.ConnectorManager != null)
                {
                    bool isConnectWC = IsFittingGroup(fmlIns, out _);

                    if (isConnectWC)
                    {
                        foreach (Connector connector in fmlIns.MEPModel.ConnectorManager.Connectors)
                        {
                            //if (connector.ConnectorType != ConnectorType.End)
                            //    continue;
                            connectors.Add(connector);
                        }
                    }
                }

                if (connectors.Count > 0)
                {
                    foreach (Connector cnt1 in connectors)
                    {
                        foreach (Connector cnt2 in cnt1.AllRefs)
                        {
                            Element eleCheck = cnt2.Owner;
                            if (null != eleCheck && (eleCheck is MEPCurve || eleCheck is FamilyInstance))
                            {
                                if (rmvElement_1 != null && rmvElement_1.Any(item => item.Id == eleCheck.Id))
                                    continue;
                                else
                                {
                                    components.Add(eleCheck);
                                    List<Element> rmvElement_2 = new List<Element>(rmvElement_1);
                                    rmvElement_2.Add(eleCheck);

                                    components.AddRange(GetAllElementConnected(document, eleCheck.Id, false, rmvElement_2));
                                }
                            }
                        }
                    }
                }
            }

            return components;
        }

        private static MEPCurve CreateNewMepCurve(Document doc, MEPCurve mepCurve1, MEPCurve mepCurve2)
        {
            MEPCurve newmEPCurve = null;
            if (mepCurve1 == null || !mepCurve1.IsValidObject || mepCurve2 == null || !mepCurve2.IsValidObject)
                return newmEPCurve;
            try
            {
                Line lineMain = ((LocationCurve)mepCurve1.Location).Curve as Line;
                Line lineBranch = ((LocationCurve)mepCurve2.Location).Curve as Line;
                //if (!Common.IsParallel(lineMain.Direction, lineBranch.Direction))
                //    return newmEPCurve;
                doc.Regenerate();
                ConnectorUtils.GetConnectorOppositeFurthestClosedTo(mepCurve1.ConnectorManager, mepCurve2.ConnectorManager, out Connector con1, out Connector con2);
                if (con1 != null && con2 != null)
                {
                    FamilyInstance fitting1 = ConnectorUtils.GetElementConnectedWithConnector(con1) as FamilyInstance;
                    FamilyInstance fitting2 = ConnectorUtils.GetElementConnectedWithConnector(con2) as FamilyInstance;

                    XYZ start = con1.Origin;
                    XYZ end = con2.Origin;

                    ICollection<ElementId> elementIds = ElementTransformUtils.CopyElement(doc, mepCurve1.Id, XYZ.Zero);

                    newmEPCurve = doc.GetElement(elementIds.FirstOrDefault()) as MEPCurve;

                    ((LocationCurve)newmEPCurve.Location).Curve = Line.CreateBound(start, end);

                    doc.Delete(new List<ElementId> { mepCurve1.Id, mepCurve2.Id });
                    doc.Regenerate();

                    if (newmEPCurve != null && fitting1 != null)
                    {
                        ConnectorUtils.GetConnectorClosedTo(newmEPCurve.ConnectorManager, fitting1.MEPModel?.ConnectorManager, out Connector con11, out Connector con22);
                        if (con11 != null && con22 != null && !con11.IsConnectedTo(con22))
                            con11.ConnectTo(con22);
                    }
                    if (newmEPCurve != null && fitting2 != null)
                    {
                        ConnectorUtils.GetConnectorClosedTo(newmEPCurve.ConnectorManager, fitting2.MEPModel?.ConnectorManager, out Connector con11, out Connector con22);
                        if (con11 != null && con22 != null && !con11.IsConnectedTo(con22))
                            con11.ConnectTo(con22);
                    }
                }
            }
            catch (Exception)
            {
            }
            return newmEPCurve;
        }

        private static List<Pipe> PickPipes(UIDocument uidoc, out Line lineMEP, string promt = "Select Pipes:")
        {
            lineMEP = null;
            List<Pipe> pipes = new List<Pipe>();
            try
            {
                var lstPipes = uidoc.Selection.PickObjects(ObjectType.Element, new PipeSelectionFilter(), promt).Select(x => uidoc.Document.GetElement(x))
                     .Where(x => x is Pipe).Cast<Pipe>().ToList();

                List<XYZ> pointPickeds = new List<XYZ>();

                lstPipes = lstPipes.OrderBy(x => CenterPipe(x).X).ThenBy(x => CenterPipe(x).Y).ToList();

                foreach (var item in lstPipes)
                {
                    pipes.Add(item);

                    Line line = (item.Location as LocationCurve).Curve as Line;

                    pointPickeds.Add(line.GetEndPoint(0));
                    pointPickeds.Add(line.GetEndPoint(1));
                }

                pointPickeds = pointPickeds.OrderBy(x => x.X).ThenBy(x => x.Y).ToList();

                XYZ p1 = null;
                XYZ p2 = null;
                double maxDistSquared = 0.0;

                int count = pointPickeds.Count;

                // 2. Duyệt qua từng cặp điểm
                for (int i = 0; i < count; i++)
                {
                    // j bắt đầu từ i + 1 để không lặp lại cặp đã kiểm tra và không so sánh với chính nó
                    for (int j = i + 1; j < count; j++)
                    {
                        XYZ vector = pointPickeds[i] - pointPickeds[j];

                        // Lấy độ dài bình phương của vector (nhanh hơn lấy độ dài thực tế)
                        double currentDistSquared = vector.GetLength();

                        if (currentDistSquared > maxDistSquared)
                        {
                            maxDistSquared = currentDistSquared;
                            p1 = pointPickeds[i];
                            p2 = pointPickeds[j];
                        }
                    }
                }

                List<XYZ> lstP = new List<XYZ>() { p1, p2 };
                lstP = lstP.OrderBy(x => x.X).ThenBy(x => x.Y).ToList();

                lineMEP = Line.CreateBound(lstP.FirstOrDefault(), lstP.LastOrDefault());

                //if (lineMEP.Direction.DotProduct(XYZ.BasisX) < 0)
                //    lineMEP = Line.CreateBound(p2, p1);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
            }
            return pipes;
        }

        private static bool IsPointOnPipe(Pipe pipe, XYZ point, double offset = 100 / 304.8)
        {
            if (pipe == null || point == null)
                return false;
            Line line = (pipe.Location as LocationCurve).Curve as Line;

            XYZ p = Common.GetPointProjectOnLine(line, point);

            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);

            p0 = p0 + line.Direction * offset;
            p1 = p1 - line.Direction * offset;

            XYZ vec1 = (p - p0).Normalize();
            XYZ vec2 = (p - p1).Normalize();

            if (vec1.DotProduct(vec2) < 0 || p.IsAlmostEqualTo(p1) || p.IsAlmostEqualTo(p0))
                return true;

            return false;
        }

        private static bool IsPointOnLine(Line line, XYZ point, double offset = 100 / 304.8)
        {
            if (line == null || point == null)
                return false;

            XYZ p = Common.GetPointProjectOnLine(line, point);

            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);

            p0 = p0 + line.Direction * offset;
            p1 = p1 - line.Direction * offset;

            XYZ vec1 = (p - p0).Normalize();
            XYZ vec2 = (p - p1).Normalize();

            if (vec1.DotProduct(vec2) < 0 || p.IsAlmostEqualTo(p1) || p.IsAlmostEqualTo(p0))
                return true;

            return false;
        }

        private static XYZ CenterPipe(Pipe pipe)
        {
            if (pipe == null) return null;

            LocationCurve lcCurve = pipe.Location as LocationCurve;
            if (lcCurve == null)
                return null;

            // Point center
            var centerP = (lcCurve.Curve.GetEndPoint(1) + lcCurve.Curve.GetEndPoint(0)) / 2;
            return new XYZ(centerP.X, centerP.Y, centerP.Z);
        }

        private bool IsCreateElbow(Line lineMep, XYZ point, Connector conNotConnec, double defaultLength = 500 / 304.8)
        {
            if (lineMep == null || point == null || conNotConnec == null)
                return false;

            XYZ p = Common.GetPointProjectOnLine(lineMep, point);

            XYZ vec1 = (p - lineMep.GetEndPoint(0)).Normalize();
            XYZ vec2 = (p - lineMep.GetEndPoint(1)).Normalize();

            //if (vec1.DotProduct(vec2) > 0)
            //    return true;

            double distance1 = p.DistanceTo(lineMep.GetEndPoint(0));

            double distance2 = p.DistanceTo(lineMep.GetEndPoint(1));

            double distanceCheck = p.DistanceTo(conNotConnec.Origin);// (isFirt) ? distance1 : distance2;

            if (distanceCheck < defaultLength || (Common.IsEqual(distanceCheck, defaultLength)))
                return true;

            return false;
        }

        private bool IsConnectedEnd(Pipe pipe, out Connector conNotConnec)
        {
            conNotConnec = null;
            if (pipe == null) return false;

            Connector con0 = pipe.ConnectorManager.Lookup(0);
            Connector con1 = pipe.ConnectorManager.Lookup(1);

            conNotConnec = (!con0.IsConnected) ? con0 : (!con1.IsConnected) ? con1 : null;

            return con0.IsConnected && con1.IsConnected;
        }
    }
}