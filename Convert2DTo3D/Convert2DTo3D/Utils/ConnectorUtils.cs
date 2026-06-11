using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D.Utils
{
    public static class ConnectorUtils
    {
        public static Connector GetConnectorPrimary(FamilyInstance instance, out Connector cSecond)
        {
            cSecond = null;

            List<Connector> connectors = ToList(instance.MEPModel.ConnectorManager).Where(x => x.Shape == ConnectorProfileType.Rectangular).ToList();

            if (connectors?.Count <= 0)
                return null;

            Connector cPrimary = connectors.FirstOrDefault(x => x.GetMEPConnectorInfo().IsPrimary == true);

            if (cPrimary == null)
                return null;

            cSecond = connectors.FirstOrDefault(x => Common.IsParallel(x.CoordinateSystem.BasisZ, cPrimary.CoordinateSystem.BasisZ)
            && x.Id != cPrimary.Id);

            return cPrimary;
        }

        public static Line ToLineUnbound(this Connector con)
        {
            if (con == null)
                return null;

            return Line.CreateUnbound(con.Origin, con.CoordinateSystem.BasisZ);
        }

        public static Line ToLineBound(this Connector con, double length = 1000 / 304.8)
        {
            if (con == null)
                return null;

            return Line.CreateBound(con.Origin, con.Origin + con.CoordinateSystem.BasisZ * length);
        }

        public static void DisconnectFrom(Connector conInput, out Element eleInput)
        {
            eleInput = null;

            if (conInput != null && conInput.IsConnected)
            {
                Element main = conInput.Owner as Element;

                foreach (Connector item in conInput.AllRefs)
                {
                    if (item != null && item.IsConnectedTo(conInput))
                    {
                        Element ele = item.Owner;

                        if (ele != null && ele.Id != main.Id && (ele is FamilyInstance || ele is MEPCurve))
                        {
                            if (ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctInsulations
                                && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeInsulations
                                && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctLinings)
                            {
                                eleInput = ele;
                                conInput.DisconnectFrom(item);
                                break;
                            }
                        }
                    }
                }
            }
        }

        public static List<Element> GetMepCurves(ConnectorManager connectorManager)
        {
            List<Element> pipes = new List<Element>();
            foreach (Connector con in connectorManager.Connectors)
            {
                if (con != null)
                {
                    Element ele = GetElementConnectedWithConnector(con);
                    if (ele != null)
                        pipes.Add(ele);
                }
            }

            return pipes;
        }

        public static void DisconnectFrom(FamilyInstance fittingWye, out Connector connectedSt, out Connector connectedEnd, out Element eleSt, out Element eleEnd)
        {
            connectedSt = null;
            connectedEnd = null;
            eleSt = null;
            eleEnd = null;
            if (fittingWye != null)
            {
                Common.GetInformationConectorWye(fittingWye, null, out Connector conSt, out Connector conEnd, out Connector conNhanhWye);

                if (conSt != null && conSt.IsConnected)
                {
                    foreach (Connector item in conSt.AllRefs)
                    {
                        if (item != null && item.IsConnectedTo(conSt))
                        {
                            conSt.DisconnectFrom(item);

                            if (item != null && item.Owner != null && item.Owner.Id != fittingWye.Id)
                            {
                                connectedSt = item;
                                eleSt = item.Owner;
                            }
                        }
                    }
                }

                if (conEnd != null && conEnd.IsConnected)
                {
                    foreach (Connector item in conEnd.AllRefs)
                    {
                        if (item != null && item.IsConnectedTo(conEnd))
                        {
                            conEnd.DisconnectFrom(item);

                            if (item != null && item.Owner != null && item.Owner.Id != fittingWye.Id)
                            {
                                connectedEnd = item;
                                eleEnd = item.Owner;
                            }
                        }
                    }
                }
            }
        }

        public static List<Connector> ToList(ConnectorManager connectorManager)
        {
            List<Connector> retval = new List<Connector>();

            if (connectorManager != null)
            {
                foreach (Connector con in connectorManager.Connectors)
                {
                    if (con != null)
                        retval.Add(con);
                }
            }

            return retval;
        }

        public static Connector GetConnectorConnectedWithConnector(Connector con)
        {
            if (con != null && con.IsConnected)
            {
                foreach (Connector item in con.AllRefs)
                {
                    if (item != null)
                        return item;
                }
            }

            return null;
        }

        public static Element GetElementConnectedWithConnector(Connector con)
        {
            if (con != null && con.IsConnected)
            {
                Element main = con.Owner as Element;

                foreach (Connector item in con.AllRefs)
                {
                    Element ele = item.Owner;
                    if (null != ele && main.Id != ele.Id && (ele is FamilyInstance || ele is MEPCurve))
                    {
                        if (ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctInsulations
                               && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeInsulations
                               && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctLinings)
                            return ele;
                    }
                }
            }
            return null;
        }

        public static Element GetElementConnectedWithConnector2(Connector con)
        {
            if (con != null && con.IsConnected)
            {
                Element main = con.Owner as Element;

                foreach (Connector item in con.AllRefs)
                {
                    Element ele = item.Owner;
                    if (null != ele && main.Id != ele.Id && (ele is FamilyInstance || ele is MEPCurve))
                    {
                        if (ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctInsulations
                               && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeInsulations
                               && ele.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctLinings)
                        {
                            return ele;
                        }
                    }
                }
            }
            return null;
        }

        public static void GetConnectorOppositeNearestClosedTo(ConnectorManager connectorManager1, List<Connector> connectors2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectors2 != null && connectors2.Count >= 1)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    foreach (Connector item2 in connectors2)
                    {
                        if (item1.CoordinateSystem.BasisZ.DotProduct(item2.CoordinateSystem.BasisZ) < 0)
                        {
                            double distance = item1.Origin.DistanceTo(item2.Origin);
                            if (distance < distanceMin)
                            {
                                con1 = item1;
                                con2 = item2;
                                distanceMin = distance;
                            }
                        }
                    }
                }
            }
        }

        public static void GetConnectorOppositeNearestClosedTo(List<Connector> connectors1, List<Connector> connectors2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectors1 != null && connectors2 != null && connectors1.Count >= 1 && connectors2.Count >= 1)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectors1)
                {
                    foreach (Connector item2 in connectors2)
                    {
                        if (item1.CoordinateSystem.BasisZ.DotProduct(item2.CoordinateSystem.BasisZ) < 0)
                        {
                            double distance = item1.Origin.DistanceTo(item2.Origin);
                            if (distance < distanceMin)
                            {
                                con1 = item1;
                                con2 = item2;
                                distanceMin = distance;
                            }
                        }
                    }
                }
            }
        }

        public static Connector GetConnectorNearestInPlan(XYZ point, ConnectorManager connectorManager, out Connector outFarest)
        {
            Connector retval = null;
            outFarest = null;

            if (point != null && connectorManager != null)
            {
                point = new XYZ(point.X, point.Y, 0);

                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    XYZ conPoint = new XYZ(item.Origin.X, item.Origin.Y, 0);
                    double distance = conPoint.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                    // lấy connector xa nhất
                    if (distance > min)
                    {
                        min = distance;
                        outFarest = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy connector ở vị trí thấp nhất và cao nhất
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="conHigher"></param>
        /// <returns></returns>
        public static Connector GetConnectorMinZ(ConnectorManager connectorManager, out Connector conHigher)
        {
            Connector retval = null;
            conHigher = null;

            if (connectorManager != null)
            {
                double maxZ = double.MaxValue;
                double minZ = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    // lấy connector thấp nhất
                    if (item.Origin.Z < maxZ)
                    {
                        maxZ = item.Origin.Z;
                        retval = item;
                    }
                    // lấy connector cao nhất
                    if (item.Origin.Z > minZ)
                    {
                        minZ = item.Origin.Z;
                        conHigher = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy connector ở vị trí thấp nhất và cao nhất
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="conHigher"></param>
        /// <returns></returns>
        public static Connector GetConnectorMinZ(Pipe pipe, out Connector conHigher)
        {
            Connector retval = null;
            conHigher = null;

            if (pipe != null)
            {
                ConnectorManager connectorManager = pipe.ConnectorManager;

                double maxZ = double.MaxValue;
                double minZ = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    // lấy connector thấp nhất
                    if (item.Origin.Z < maxZ)
                    {
                        maxZ = item.Origin.Z;
                        retval = item;
                    }
                    // lấy connector cao nhất
                    if (item.Origin.Z > minZ)
                    {
                        minZ = item.Origin.Z;
                        conHigher = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy connector chưa được kết nối
        /// </summary>
        /// <param name="connectorManager"></param>
        /// <returns></returns>
        public static Connector GetConnectorNotConnnected(ConnectorManager connectorManager)
        {
            if (connectorManager != null)
            {
                foreach (Connector con in connectorManager.Connectors)
                {
                    if (!con.IsConnected)
                        return con;
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy connector  đã kết nối
        /// </summary>
        /// <param name="connectorManager"></param>
        /// <returns></returns>
        public static Connector GetConnectorConnnected(ConnectorManager connectorManager)
        {
            if (connectorManager != null)
            {
                foreach (Connector con in connectorManager.Connectors)
                {
                    if (con.IsConnected)
                        return con;
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy connector có cao độ cao hơn
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        public static Connector GetConnectorValid(Pipe pipe)
        {
            if (pipe != null && pipe.Location is LocationCurve locationCurve)
            {
                double slope = Math.Round((double)ParameterUtils.GetValueParameterByBuilt(pipe, BuiltInParameter.RBS_PIPE_SLOPE), 5);

                XYZ pointSt = locationCurve.Curve.GetEndPoint(0);

                Connector conSt = GetConnectorNearest(pointSt, pipe, out Connector conEnd);

                if (Common.IsEqual(slope, 0))
                {
                    if (conEnd.IsConnected)
                        return conSt;

                    return conEnd;
                }
                else
                {
                    Connector retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conSt : conEnd;
                    if (retval.IsConnected)
                        return (conSt.Origin.Z < conEnd.Origin.Z) ? conSt : conEnd;

                    return retval;
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy connector có cao độ cao hơn
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        public static Connector GetConnectorHigher(Pipe pipe)
        {
            if (pipe != null && pipe.Location is LocationCurve locationCurve)
            {
                double slope = Math.Round((double)ParameterUtils.GetValueParameterByBuilt(pipe, BuiltInParameter.RBS_PIPE_SLOPE), 7);

                XYZ pointSt = locationCurve.Curve.GetEndPoint(0);

                Connector conSt = GetConnectorNearest(pointSt, pipe, out Connector conEnd);

                if (Common.IsEqual(slope, 0))
                {
                    if (!conSt.IsConnected)
                        return conSt;

                    return conEnd;
                }
                else
                {
                    Connector retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conSt : conEnd;
                    if (retval.IsConnected)
                        return null;

                    return retval;
                }
            }

            return null;
        }

        public static Connector GetConnectorNearest(XYZ point, ConnectorManager connectorManager, out Connector outFarest, bool is2D = false)
        {
            Connector retval = null;
            outFarest = null;

            if (point != null && connectorManager != null)
            {
                if (is2D)
                    point = Common.To2D(point);

                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    if (item.ConnectorType != ConnectorType.End)
                        continue;

                    XYZ conPoint = new XYZ(item.Origin.X, item.Origin.Y, item.Origin.Z);

                    if (is2D)
                        conPoint = Common.To2D(conPoint);

                    double distance = conPoint.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                    // lấy connector xa nhất
                    if (distance > min)
                    {
                        min = distance;
                        outFarest = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy ra connector gần nhất và xa nhất với 1 điểm cho trước
        /// </summary>
        /// <param name="point"></param>
        /// <param name="pipe"></param>
        /// <param name="outFarest"></param>
        /// <returns></returns>
        public static Connector GetConnectorNearest(XYZ point, Pipe pipe, out Connector outFarest)
        {
            Connector retval = null;
            outFarest = null;

            if (point != null && pipe != null)
            {
                ConnectorManager connectorManager = pipe.ConnectorManager;

                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    double distance = item.Origin.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                    // lấy connector xa nhất
                    if (distance > min)
                    {
                        min = distance;
                        outFarest = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy ra connector gần nhất và xa nhất với 1 điểm cho trước
        /// </summary>
        /// <param name="point"></param>
        /// <param name="pipe"></param>
        /// <param name="outFarest"></param>
        /// <returns></returns>
        public static Connector GetConnectorNearestInXYPlan(XYZ point, Pipe pipe, out Connector outFarest)
        {
            Connector retval = null;
            outFarest = null;

            if (point != null && pipe != null)
            {
                point = new XYZ(point.X, point.Y, 0);

                ConnectorManager connectorManager = pipe.ConnectorManager;

                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    if (item == null)
                        continue;

                    XYZ originCon = new XYZ(item.Origin.X, item.Origin.Y, 0);
                    double distance = originCon.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                    // lấy connector xa nhất
                    if (distance > min)
                    {
                        min = distance;
                        outFarest = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy 2 connector thuộc 2 ống khác nhau và gần nhau nhất
        /// </summary>
        /// <param name="connectorManager1"></param>
        /// <param name="connectorManager2"></param>
        /// <param name="con1"></param>
        /// <param name="con2"></param>
        public static void GetConnectorClosedToInXYPlan(ConnectorManager connectorManager1, ConnectorManager connectorManager2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectorManager2 != null)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    foreach (Connector item2 in connectorManager2.Connectors)
                    {
                        XYZ oringin1 = new XYZ(item1.Origin.X, item1.Origin.Y, 0);
                        XYZ oringin2 = new XYZ(item2.Origin.X, item2.Origin.Y, 0);

                        double distance = oringin1.DistanceTo(oringin2);
                        if (distance < distanceMin)
                        {
                            con1 = item1;
                            con2 = item2;
                            distanceMin = distance;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Lấy 2 connector thuộc 2 ống khác nhau và gần nhau nhất
        /// </summary>
        /// <param name="connectorManager1"></param>
        /// <param name="connectorManager2"></param>
        /// <param name="con1"></param>
        /// <param name="con2"></param>
        public static void GetConnectorClosedTo(ConnectorManager connectorManager1, ConnectorManager connectorManager2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectorManager2 != null)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    foreach (Connector item2 in connectorManager2.Connectors)
                    {
                        double distance = item1.Origin.DistanceTo(item2.Origin);
                        if (distance < distanceMin)
                        {
                            con1 = item1;
                            con2 = item2;
                            distanceMin = distance;
                        }
                    }
                }
            }
        }

        public static Connector GetConnectorNearestClosedTo(Connector conInput, ConnectorManager connectorManager, bool isOpposite = true)
        {
            Connector retval = null;

            if (connectorManager != null)

            {
                double distanceMax = double.MaxValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    if (item.ConnectorType != ConnectorType.End)
                        continue;

                    bool isErr = false;
                    if (isOpposite)
                        isErr = conInput.CoordinateSystem.BasisZ.DotProduct(item.CoordinateSystem.BasisZ) < 0;
                    else
                        isErr = conInput.CoordinateSystem.BasisZ.DotProduct(item.CoordinateSystem.BasisZ) > 0;

                    if (isErr)
                    {
                        double distance = conInput.Origin.DistanceTo(item.Origin);
                        if (distance < distanceMax)
                        {
                            retval = item;
                            distanceMax = distance;
                        }
                    }
                }
            }
            return retval;
        }

        public static Connector GetConnectorFurthestClosedTo(Connector conInput, ConnectorManager connectorManager, bool isOpposite = true)
        {
            Connector retval = null;

            if (connectorManager != null)

            {
                double distanceMax = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    if (item.ConnectorType != ConnectorType.End)
                        continue;

                    bool isErr = false;
                    if (isOpposite)
                        isErr = conInput.CoordinateSystem.BasisZ.DotProduct(item.CoordinateSystem.BasisZ) < 0;
                    else
                        isErr = conInput.CoordinateSystem.BasisZ.DotProduct(item.CoordinateSystem.BasisZ) > 0;

                    if (isErr)
                    {
                        double distance = conInput.Origin.DistanceTo(item.Origin);
                        if (distance > distanceMax)
                        {
                            retval = item;
                            distanceMax = distance;
                        }
                    }
                }
            }
            return retval;
        }

        public static void GetConnectorOppositeNearestClosedTo(ConnectorManager connectorManager1, ConnectorManager connectorManager2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectorManager2 != null)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    if (item1.ConnectorType != ConnectorType.End)
                        continue;

                    foreach (Connector item2 in connectorManager2.Connectors)
                    {
                        if (item2.ConnectorType != ConnectorType.End)
                            continue;

                        if (item1.CoordinateSystem.BasisZ.DotProduct(item2.CoordinateSystem.BasisZ) < 0)
                        {
                            double distance = item1.Origin.DistanceTo(item2.Origin);
                            if (distance < distanceMin)
                            {
                                con1 = item1;
                                con2 = item2;
                                distanceMin = distance;
                            }
                        }
                    }
                }
            }
        }

        public static void GetConnectorOppositeFurthestClosedTo(ConnectorManager connectorManager1, ConnectorManager connectorManager2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectorManager2 != null)

            {
                double distanceMax = double.MinValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    if (item1.ConnectorType != ConnectorType.End)
                        continue;

                    foreach (Connector item2 in connectorManager2.Connectors)
                    {
                        if (item2.ConnectorType != ConnectorType.End)
                            continue;

                        if (item1.CoordinateSystem.BasisZ.DotProduct(item2.CoordinateSystem.BasisZ) < 0)
                        {
                            double distance = item1.Origin.DistanceTo(item2.Origin);
                            if (distance > distanceMax)
                            {
                                con1 = item1;
                                con2 = item2;
                                distanceMax = distance;
                            }
                        }
                    }
                }
            }
        }
    }
}