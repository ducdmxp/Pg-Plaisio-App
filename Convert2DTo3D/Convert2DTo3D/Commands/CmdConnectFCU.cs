using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using CheckPanelProject.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using ParameterUtils = Convert2DTo3D.Utils.ParameterUtils;

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdConnectFCU : IExternalCommand
    {
        public const string Convert2DTo3D_ConnectedFCU = "Convert2DTo3D_ConnectedFCU";

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

            ConnectFCUFrm form = new ConnectFCUFrm();

            if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return Result.Cancelled;

            TransactionGroup tranG = new TransactionGroup(doc, "ConnectFCUs");

            try
            {
                List<FamilyInstance> lstFCUs = uiDoc.Selection.PickObjects(ObjectType.Element,
               new FamilyinstanSelectionFilterCategory(new List<BuiltInCategory> { BuiltInCategory.OST_MechanicalEquipment }), "Please select valve :")
                      .Select(x => doc.GetElement(x) as FamilyInstance).ToList();

                tranG.Start();

                foreach (var instanceFCU in lstFCUs)
                {
                    Connector cPrimary = ConnectorUtils.GetConnectorPrimary(instanceFCU, out Connector cSecond);

                    if (cPrimary == null || cSecond == null)
                        continue;

                    if (form.ConnectionMode == 0)
                    {
                        ConnectFCU(doc, instanceFCU, cPrimary, form.SymbolSimiliIdInput,
                            form.SymboHopGioIdInput, form.SystemTypeIdInput, form.DuctTypeIdInput,
                            form.LenghtInput / 304.8, form.HeightInput / 304.8, form.WidthInput / 304.8, form.TypeConnectInput);
                    }
                    else if (form.ConnectionMode == 1)
                    {
                        ConnectFCU(doc, instanceFCU, cSecond, form.SymbolSimiliIdOutput,
                            form.SymboHopGioIdOutput, form.SystemTypeIdOutput, form.DuctTypeIdOutput,
                            form.LenghtOutput / 304.8, form.HeightOutput / 304.8, form.WidthOutput / 304.8, form.TypeConnectOutput);
                    }
                    else
                    {
                        ConnectFCU(doc, instanceFCU, cPrimary, form.SymbolSimiliIdInput,
                            form.SymboHopGioIdInput, form.SystemTypeIdInput, form.DuctTypeIdInput,
                            form.LenghtInput / 304.8, form.HeightInput / 304.8, form.WidthInput / 304.8, form.TypeConnectInput);

                        ConnectFCU(doc, instanceFCU, cSecond, form.SymbolSimiliIdOutput,
                            form.SymboHopGioIdOutput, form.SystemTypeIdOutput, form.DuctTypeIdOutput,
                            form.LenghtOutput / 304.8, form.HeightOutput / 304.8, form.WidthOutput / 304.8, form.TypeConnectOutput);
                    }
                }

                tranG.Assimilate();
            }
            catch (Exception ex)
            {
                if (tranG.HasStarted())
                    tranG.RollBack();

                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            return Result.Succeeded;
        }

        private bool ConnectFCU(Document doc, FamilyInstance instanceFCU, Connector connector,
            FamilySymbol symbolSimili, FamilySymbol symbolHopGio,
            ElementId mepSystemId, ElementId ductTypeId,
            double lenght, double height, double width,
            int typeConnect)
        {
            Transaction tran = new Transaction(doc, "ConnectFCU");

            try
            {
                List<Element> lstElementCreateds = new List<Element>();

                Level level = doc.GetElement(instanceFCU.LevelId) as Level;
                if (level == null)
                    return false;

                XYZ lcBefore = instanceFCU.GetLocationPoint();

                tran.Start();

                if (!symbolSimili.IsActive)
                    symbolSimili.Activate();

                if (!symbolHopGio.IsActive)
                    symbolHopGio.Activate();

                XYZ location1 = connector.Origin + connector.CoordinateSystem.BasisZ * 100 / 304.8;

                FamilyInstance simili = doc.Create.NewFamilyInstance(location1, symbolSimili, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                ParameterUtils.SetValueParameterByBuiltIn(simili, BuiltInParameter.INSTANCE_ELEVATION_PARAM, location1.Z - level.Elevation);

                doc.Regenerate();

                lstElementCreateds.Add(simili);

                Connector cSimili1 = ConnectorUtils.GetConnectorNearest(connector.Origin, simili.MEPModel.ConnectorManager, out Connector cSimili2);

                Common.RotateLine(doc, simili, connector.ToLineUnbound(), cSimili1.ToLineBound());

                doc.Regenerate();

                cSimili1 = ConnectorUtils.GetConnectorNearest(connector.Origin, simili.MEPModel.ConnectorManager, out cSimili2);

                cSimili1.Height = connector.Height;
                cSimili1.Width = connector.Width;

                connector.ConnectTo(cSimili1);

                doc.Regenerate();

                XYZ p0 = cSimili2.Origin + XYZ.BasisZ * (height - cSimili1.Height) / 2;

                XYZ p1 = p0 + cSimili2.CoordinateSystem.BasisZ * lenght;

                Duct duct = Duct.Create(doc, mepSystemId, ductTypeId, level.Id, p0, p1);

                lstElementCreateds.Add(duct);

                if (typeConnect == 1)
                {
                    ParameterUtils.SetValueParameterByBuiltIn(duct, BuiltInParameter.RBS_CURVE_WIDTH_PARAM, width);

                    ParameterUtils.SetValueParameterByBuiltIn(duct, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, height);
                }

                doc.Regenerate();

                tran.Commit();

                if (duct != null && simili != null)
                {
                    tran.Start();

                    ConnectorUtils.GetConnectorClosedTo(duct.ConnectorManager, simili.MEPModel.ConnectorManager, out Connector con01, out Connector con02);

                    if (con01 != null && con02 != null)
                    {
                        FamilyInstance fittingTran = doc.Create.NewTransitionFitting(con01, con02);

                        lstElementCreateds.Add(fittingTran);

                        if (typeConnect == 0 && fittingTran != null)
                            doc.Delete(fittingTran.Id);
                    }

                    tran.Commit();

                    tran.Start();

                    ((LocationCurve)duct.Location).Curve = Line.CreateBound(p0, p0 + (cSimili2.CoordinateSystem.BasisZ * lenght));

                    tran.Commit();
                }

                //kiểu kết nối
                if (typeConnect == 0 && duct != null)
                {
                    tran.Start();

                    ConnectorUtils.GetConnectorClosedTo(duct.ConnectorManager, simili.MEPModel.ConnectorManager, out Connector con01, out Connector con02);

                    if (con01 != null && con02 != null && con01.IsConnectedTo(con02) == false)
                    {
                        con01.Height = cSimili2.Height;
                        con01.Width = cSimili2.Width;

                        con01.ConnectTo(con02);

                        doc.Regenerate();
                    }

                    tran.Commit();

                    tran.Start();

                    doc.Delete(duct.Id);

                    XYZ location2 = cSimili2.Origin + cSimili2.CoordinateSystem.BasisZ * 100 / 304.8;

                    FamilyInstance hopGioCap = doc.Create.NewFamilyInstance(location2, symbolHopGio, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    ParameterUtils.SetValueParameterByBuiltIn(hopGioCap, BuiltInParameter.INSTANCE_ELEVATION_PARAM, location2.Z - level.Elevation);

                    doc.Regenerate();

                    lstElementCreateds.Add(hopGioCap);

                    Connector cHopGioCap1 = ConnectorUtils.ToList(hopGioCap.MEPModel.ConnectorManager).FirstOrDefault(x => x.Shape == ConnectorProfileType.Rectangular);

                    Common.RotateLine(doc, hopGioCap, connector.ToLineUnbound(), cHopGioCap1.ToLineBound());

                    doc.Regenerate();

                    cHopGioCap1 = ConnectorUtils.ToList(hopGioCap.MEPModel.ConnectorManager).FirstOrDefault(x => x.Shape == ConnectorProfileType.Rectangular);

                    cHopGioCap1.Height = cSimili2.Height;

                    cHopGioCap1.Width = cSimili2.Width;

                    if (cHopGioCap1.CoordinateSystem.BasisZ.DotProduct(cSimili2.CoordinateSystem.BasisZ) > 0)
                    {
                        ElementTransformUtils.RotateElement(doc, hopGioCap.Id, Line.CreateUnbound(cHopGioCap1.Origin, hopGioCap.FacingOrientation), Math.PI);
                    }

                    cHopGioCap1.ConnectTo(cSimili2);

                    doc.Regenerate();

                    tran.Commit();
                }

                tran.Start();

                List<ElementId> lstElementCreatedIds = lstElementCreateds.Where(x => x != null && x.IsValidObject).Select(x => x.Id).ToList();

                if (typeConnect == 0)
                    lstElementCreatedIds = new List<ElementId> { instanceFCU.Id };

                XYZ vector = XYZ.BasisZ * 1 / 304.8;

                if (duct != null && duct.IsValidObject)
                {
                    ElementTransformUtils.MoveElement(doc, duct.Id, vector);

                    doc.Regenerate();

                    ElementTransformUtils.MoveElement(doc, duct.Id, vector.Negate());

                    doc.Regenerate();
                }

                ElementTransformUtils.MoveElements(doc, lstElementCreatedIds, lcBefore - instanceFCU.GetLocationPoint());

                ElementTransformUtils.MoveElements(doc, lstElementCreatedIds, vector);

                doc.Regenerate();

                ElementTransformUtils.MoveElements(doc, lstElementCreatedIds, vector.Negate());

                doc.Regenerate();

                foreach (var el in lstElementCreateds)
                {
                    if (el != null && el.IsValidObject)
                        el.SetElementParameterDataStorage<string>(Convert2DTo3D_ConnectedFCU, Convert2DTo3D_ConnectedFCU, true.ToString());
                }

                instanceFCU.SetElementParameterDataStorage<string>(Convert2DTo3D_ConnectedFCU, Convert2DTo3D_ConnectedFCU, false.ToString());

                tran.Commit();
            }
            catch (Exception ex)
            {
                if (tran.HasStarted())
                    tran.RollBack();

                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            return true;
        }

        private double GetLenghtDuct(DuctType ductType)
        {
            Transaction tran = new Transaction(ductType.Document, "GetLenghtDuct");

            try
            {
                tran.Start();

                FamilySymbol symbol = Common.GetSymbolSeted(ductType.Document, ductType, RoutingPreferenceRuleGroupType.Transitions);

                if (symbol != null)
                {
                    if (!symbol.IsActive)
                        symbol.Activate();

                    FamilyInstance instance = ductType.Document.Create.NewFamilyInstance(XYZ.Zero, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    ductType.Document.Regenerate();

                    if (instance != null)
                    {
                        object objLenght = ParameterUtils.GetValueParameterByName(instance, "Duct Length");

                        if (objLenght is double)
                            return (double)objLenght;
                        else
                        {
                            List<Connector> connectors = ConnectorUtils.ToList(instance.MEPModel.ConnectorManager);

                            if (connectors.Count >= 2)
                                return connectors[0].Origin.DistanceTo(connectors[1].Origin);
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                tran.RollBack();
            }

            return 0.0;
        }

        public static void DeleteConnectFCUBySelection(UIDocument uiDoc)
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

                elements = elements.GroupBy(x => x.Id).Select(x => x.FirstOrDefault()).Where(x => IsFittingGroup(x, out _)).ToList();

                List<Element> lstMEPDeletes = elements.Where(x => IsFittingGroup(x, out string value) && value.Equals(true.ToString())).ToList();

                uiDoc.Document.Delete(lstMEPDeletes.Select(x => x.Id).ToList());
                uiDoc.Document.Regenerate();

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
                string valueEntity = element.GetElementParameterDataStorage<string>(Convert2DTo3D_ConnectedFCU, Convert2DTo3D_ConnectedFCU);

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
    }

    public class FamilyinstanSelectionFilterCategory : ISelectionFilter
    {
        private List<BuiltInCategory> ListCategoryId;

        public FamilyinstanSelectionFilterCategory(List<BuiltInCategory> listCategoryId = null)
        {
            ListCategoryId = listCategoryId;
        }

        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance instance)
            {
                if (instance.MEPModel != null && instance.MEPModel.ConnectorManager != null)
                {
                    if (ListCategoryId == null || ListCategoryId.Count == 0 || ListCategoryId.Any(x => ((int)x) == instance.Category.Id.IntegerValue))
                        return true;
                }
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ point)
        {
            return true;
        }
    }
}