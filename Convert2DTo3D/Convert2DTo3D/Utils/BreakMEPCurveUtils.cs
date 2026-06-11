using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D.Utils
{
    public class BreakMEPCurveUtils
    {
        /// <summary>
        /// SplitMEPCurve
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="mEPCurve"></param>
        /// <param name="fittingWye"></param>
        /// <returns></returns>
        public static bool SplitMEPCurve(Document doc, MEPCurve mEPCurve, FamilyInstance fittingWye, out MEPCurve splitMEPCurve)
        {
            splitMEPCurve = null;
            try
            {
                if (mEPCurve != null && fittingWye != null)
                {
                    XYZ lcBefore = (fittingWye.Location as LocationPoint).Point;

                    Common.GetInformationConectorWye(fittingWye, null, out Connector main1, out Connector main2, out Connector conTee);

                    Line lineFitting = Line.CreateBound(main1.Origin, main2.Origin);
                    List<MEPCurve> lstMEPCurveSplits = BreakMEPCurveToListMEPCurve(doc, mEPCurve, fittingWye);

                    foreach (var item in lstMEPCurveSplits)
                    {
                        if (item != null && item.IsValidObject)
                        {
                            XYZ centerMEPCurve = ((LocationCurve)item.Location).Curve.Evaluate(0.5, true);

                            if (Common.IsBetweenLine(lineFitting, centerMEPCurve))
                                doc.Delete(item.Id);
                        }
                    }

                    doc.Regenerate();

                    XYZ lcAffter = (fittingWye.Location as LocationPoint).Point;

                    ElementTransformUtils.MoveElement(doc, fittingWye.Id, lcBefore - lcAffter);

                    doc.Regenerate();

                    MEPCurve mepCurve1 = lstMEPCurveSplits.FirstOrDefault(x => x != null && x.IsValidObject);
                    MEPCurve mepCurve2 = lstMEPCurveSplits.LastOrDefault(x => x != null && x.IsValidObject);

                    if (mepCurve1 != null)
                    {
                        ConnectorUtils.GetConnectorClosedTo(fittingWye.MEPModel.ConnectorManager, mepCurve1.ConnectorManager, out Connector con1, out Connector con2);
                        if (con1 != null && con2 != null && !con1.IsConnectedTo(con2))
                            con2.ConnectTo(con1);
                    }

                    if (mepCurve2 != null)
                    {
                        ConnectorUtils.GetConnectorClosedTo(fittingWye.MEPModel.ConnectorManager, mepCurve2.ConnectorManager, out Connector con1, out Connector con2);
                        if (con1 != null && con2 != null && !con1.IsConnectedTo(con2))
                            con2.ConnectTo(con1);
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// BreakMEPCurve
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="mEPCurve"></param>
        /// <param name="fittingWye"></param>
        /// <returns></returns>
        public static List<MEPCurve> BreakMEPCurveToListMEPCurve(Document doc, MEPCurve mEPCurve, FamilyInstance fittingWye)
        {
            List<MEPCurve> retval = new List<MEPCurve>();
            if (mEPCurve != null && fittingWye != null)
            {
                Common.GetInformationConectorWye(fittingWye, null, out Connector conFittingStart, out Connector conFittingEnd, out Connector conNhanhWye);

                XYZ stFitting = conFittingStart.Origin;
                XYZ endFitting = conFittingEnd.Origin;

                Line line = ((LocationCurve)mEPCurve.Location).Curve as Line;

                if (mEPCurve != null && mEPCurve.IsValidObject && !retval.Select(x => x.Id).Contains(mEPCurve.Id))
                    retval.Add(mEPCurve);

                if (Common.IsBetweenLine(line, stFitting))
                {
                    ElementId newMEPCurveId = BreakMEPCurve(doc, mEPCurve.Id, stFitting);
                    if (newMEPCurveId != ElementId.InvalidElementId)
                    {
                        MEPCurve mEPCurveSplit = doc.GetElement(newMEPCurveId) as MEPCurve;
                        if (mEPCurveSplit != null && mEPCurveSplit.IsValidObject && !retval.Select(x => x.Id).Contains(mEPCurveSplit.Id))
                            retval.AddRange(BreakMEPCurveToListMEPCurve(doc, mEPCurveSplit, fittingWye));
                    }
                }

                if (Common.IsBetweenLine(line, endFitting))
                {
                    ElementId newMEPCurveId = BreakMEPCurve(doc, mEPCurve.Id, endFitting);
                    if (newMEPCurveId != ElementId.InvalidElementId)
                    {
                        MEPCurve mEPCurveSplit = doc.GetElement(newMEPCurveId) as MEPCurve;
                        if (mEPCurveSplit != null && mEPCurveSplit.IsValidObject && !retval.Select(x => x.Id).Contains(mEPCurveSplit.Id))
                            retval.AddRange(BreakMEPCurveToListMEPCurve(doc, mEPCurveSplit, fittingWye));
                    }
                }
            }

            return retval;
        }

        public static ElementId BreakMEPCurve(Document doc, ElementId mepCurveId, XYZ breakPoint)
        {
            ElementId newMEPCurveId = ElementId.InvalidElementId;
            try
            {
                if (doc != null && mepCurveId != ElementId.InvalidElementId && breakPoint != null)
                {
                    MEPCurve mepCurve = doc.GetElement(mepCurveId) as MEPCurve;

                    if (mepCurve != null && mepCurve.Location is LocationCurve lc)
                    {
                        Line line = lc.Curve as Line;

                        XYZ project = Common.GetPointProjectOnLine(line, breakPoint);

                        if (!Common.IsBetweenLine(line, project))
                            return newMEPCurveId;

                        if (mepCurve is Autodesk.Revit.DB.Mechanical.Duct duct)
                        {
                            newMEPCurveId = MechanicalUtils.BreakCurve(doc, duct.Id, project);
                        }
                        else if (mepCurve is Pipe pipe)
                        {
                            newMEPCurveId = PlumbingUtils.BreakCurve(doc, pipe.Id, project);
                        }
                        else
                        {
                            //copy mepCurveToOptimize as newPipe and move to brkPoint

                            var start = line.GetEndPoint(0);
                            var end = line.GetEndPoint(1);

                            Connector con1 = ConnectorUtils.GetConnectorNearest(start, mepCurve.ConnectorManager, out Connector con2);

                            ConnectorUtils.DisconnectFrom(con1, out Element fitting1);
                            ConnectorUtils.DisconnectFrom(con2, out Element fitting2);

                            var copiedEls = ElementTransformUtils.CopyElement(doc, mepCurve.Id, breakPoint - start);

                            newMEPCurveId = copiedEls.First();

                            MEPCurve newCableTray = doc.GetElement(newMEPCurveId) as MEPCurve;

                            if (!start.IsAlmostEqualTo(breakPoint))
                            {
                                ((LocationCurve)mepCurve.Location).Curve = Line.CreateBound(start, breakPoint);
                                if (mepCurve != null && fitting1 != null && fitting1 is FamilyInstance instance1)
                                {
                                    ConnectorUtils.GetConnectorClosedTo(mepCurve.ConnectorManager, instance1.MEPModel?.ConnectorManager, out Connector con11, out Connector con22);
                                    if (con11 != null && con22 != null && !con11.IsConnected && !con22.IsConnected)
                                        con11.ConnectTo(con22);
                                }
                            }

                            if (!end.IsAlmostEqualTo(breakPoint))
                            {
                                ((LocationCurve)newCableTray.Location).Curve = Line.CreateBound(breakPoint, end);

                                if (newCableTray != null && fitting2 != null && fitting2 is FamilyInstance instance2)
                                {
                                    ConnectorUtils.GetConnectorClosedTo(newCableTray.ConnectorManager, instance2.MEPModel?.ConnectorManager, out Connector con11, out Connector con22);
                                    if (con11 != null && con22 != null && !con11.IsConnected && !con22.IsConnected)
                                        con11.ConnectTo(con22);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                return ElementId.InvalidElementId;
            }

            return newMEPCurveId;
        }
    }
}