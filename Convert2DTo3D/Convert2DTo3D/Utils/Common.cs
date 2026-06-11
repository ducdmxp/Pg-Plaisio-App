using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;

namespace Convert2DTo3D.Utils
{
    public static class Common
    {
        #region Basic Type Conversions

        /// <summary>
        /// Converts string to double with fallback to default value
        /// </summary>
        public static double ToDouble(this string strValue, double defaultValue = 0.0)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return double.TryParse(strValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : defaultValue;
        }

        /// <summary>
        /// Converts string to int with fallback to default value
        /// </summary>
        public static int ToInt(this string strValue, int defaultValue = 0)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return int.TryParse(strValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
                ? value
                : defaultValue;
        }

        /// <summary>
        /// Converts string to decimal with fallback to default value
        /// </summary>
        public static decimal ToDecimal(this string strValue, decimal defaultValue = 0m)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return decimal.TryParse(strValue, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal value)
                ? value
                : defaultValue;
        }

        /// <summary>
        /// Converts string to long with fallback to default value
        /// </summary>
        public static long ToLong(this string strValue, long defaultValue = 0L)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return long.TryParse(strValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value)
                ? value
                : defaultValue;
        }

        /// <summary>
        /// Converts string to bool with fallback to default value
        /// </summary>
        public static bool ToBool(this string strValue, bool defaultValue = false)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return bool.TryParse(strValue, out bool value) ? value : defaultValue;
        }

        /// <summary>
        /// Converts string to DateTime with fallback to default value
        /// </summary>
        public static DateTime ToDateTime(this string strValue, DateTime defaultValue = default)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return defaultValue;

            return DateTime.TryParse(strValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime value)
                ? value
                : defaultValue;
        }

        #endregion Basic Type Conversions

        public static XYZ GetLocationPoint(this FamilyInstance instance)
        {
            if (instance != null && instance.Location is LocationPoint lcp)
                return lcp.Point;

            return null;
        }

        public static Line GetLine(this FamilyInstance instance, bool isHandOrientation = true, double lenght = 1000 / 304.8)
        {
            if (instance != null)
            {
                XYZ point = instance.GetLocationPoint();

                if (isHandOrientation)
                    return Line.CreateBound(point, point + instance.HandOrientation * lenght);
                else
                    return Line.CreateBound(point, point + instance.FacingOrientation * lenght);
            }

            return null;
        }

        public static DirectShape DrawSolid(Document doc, Solid solid)
        {
            DirectShape ds = null;

            if (doc != null && solid is GeometryObject geometry && solid.Volume > 0)
            {
                List<GeometryObject> lstGeo = new List<GeometryObject>
                {
                    geometry
                };

                var category = Category.GetCategory(doc, BuiltInCategory.OST_GenericModel);
                if (category != null)
                {
                    var dsType = DirectShapeType.Create(doc, "DrawSolid", category.Id);
                    if (dsType != null)
                    {
                        ds = DirectShape.CreateElement(doc, category.Id);

                        if (ds != null)
                        {
                            // Set type
                            ds.SetTypeId(dsType.Id);

                            // Set shape
                            ds.SetShape(lstGeo);

                            return ds;
                        }
                    }
                }
            }

            return ds;
        }

        public static XYZ To2D(this XYZ point, double z = 0.0)
        {
            if (point == null)
                return null;
            return new XYZ(point.X, point.Y, z);
        }

        public static Plane ToPlane(this PlanarFace planar)
        {
            if (planar == null)
                return null;
            return Plane.CreateByNormalAndOrigin(planar.FaceNormal, planar.Origin);
        }

        public static void RotateLine(Document doc, FamilyInstance fitting, Line axisSource, Line axisDestination)
        {
            if (doc == null || fitting == null || axisDestination == null || axisSource == null)
                return;

            if (IsParallel(axisSource.Direction, axisDestination.Direction))
                return;

            XYZ vector = axisDestination.Direction.CrossProduct(axisSource.Direction);
            XYZ intersection = GetUnBoundIntersection(axisDestination, axisSource);

            if (intersection != null)
            {
                double angle = axisDestination.Direction.AngleTo(axisSource.Direction);

                Line line = Line.CreateUnbound(intersection, vector);

                ElementTransformUtils.RotateElement(doc, fitting.Id, line, angle);
                doc.Regenerate();
            }
            else
            {
                intersection = (axisDestination.GetEndPoint(0) + axisDestination.GetEndPoint(1)) / 2;
                double angle = axisDestination.Direction.AngleTo(axisSource.Direction);

                Line line = Line.CreateUnbound(intersection, vector);

                ElementTransformUtils.RotateElement(doc, fitting.Id, line, angle);
                doc.Regenerate();
            }
        }

        public static FamilySymbol GetSymbolSeted(Document doc, DuctType ductType, RoutingPreferenceRuleGroupType checkType)
        {
            try
            {
                if (doc != null && ductType != null && ductType.IsValidObject)
                {
                    RoutingPreferenceManager rpm = ductType.RoutingPreferenceManager;

                    if (checkType == RoutingPreferenceRuleGroupType.Junctions &&
                      rpm.PreferredJunctionType != PreferredJunctionType.Tee)
                        return null;

                    int numberOfRule = rpm.GetNumberOfRules(checkType);

                    if (numberOfRule > 0)
                    {
                        for (int i = 0; i < numberOfRule; i++)
                        {
                            RoutingPreferenceRule rule = rpm.GetRule(checkType, i);

                            if (rule.MEPPartId != null &&
                                rule.MEPPartId != ElementId.InvalidElementId)
                            {
                                if (rule.NumberOfCriteria > 0)
                                {
                                    PrimarySizeCriterion primarySizeCriterion = rule.GetCriterion(0) as PrimarySizeCriterion;

                                    if (primarySizeCriterion != null)
                                        return doc.GetElement(rule.MEPPartId) as FamilySymbol;
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Check file in using
        /// </summary>
        public static bool IsFileInUse(string path)
        {
            if (File.Exists(path))
            {
                FileStream stream = null;
                try
                {
                    FileInfo file = new FileInfo(path);
                    stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                catch (IOException)
                {
                    return true;
                }
                finally
                {
                    if (stream != null)
                    {
                        stream.Close();
                    }
                }
            }

            return false;
        }

        public static void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
        {
            GetInformationConectorWye(wye, null, out Connector connector2, out Connector connector3, out Connector conTee);

            Line rotateLine = Line.CreateBound(connector2.Origin, connector3.Origin);

            if (IsParallel(axisLine.Direction, rotateLine.Direction))
                return;

            XYZ vector = rotateLine.Direction.CrossProduct(axisLine.Direction);
            XYZ intersection = GetUnBoundIntersection(rotateLine, axisLine);

            if (intersection != null)
            {
                double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

                Line line = Line.CreateUnbound(intersection, vector);

                ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
                doc.Regenerate();
            }
            else
            {
                intersection = (connector2.Origin + connector3.Origin) / 2;
                double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

                Line line = Line.CreateUnbound(intersection, vector);

                ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
                doc.Regenerate();
            }
        }

        public static XYZ LineIntersection(Curve line1, Curve line2, bool isUnbound = false)
        {
            if (line1 != null && line2 != null)
            {
                Line lineCopy1 = line1.Clone() as Line;
                Line lineCopy2 = line2.Clone() as Line;

                if (isUnbound)
                {
                    lineCopy1.MakeUnbound();
                    lineCopy1.MakeUnbound();
                }

                SetComparisonResult setComparisonResult = lineCopy1.Intersect(lineCopy2, out IntersectionResultArray iResult);
                if (setComparisonResult != SetComparisonResult.Disjoint)
                    return iResult.get_Item(0).XYZPoint;
            }

            return null;
        }

        public static XYZ GetUnBoundIntersection(Line Line1, Line Line2)
        {
            if (Line1 != null && Line2 != null)
            {
                Curve ExtendedLine1 = Line.CreateUnbound(Line1.Origin, Line1.Direction);
                Curve ExtendedLine2 = Line.CreateUnbound(Line2.Origin, Line2.Direction);
                SetComparisonResult setComparisonResult = ExtendedLine1.Intersect(ExtendedLine2, out IntersectionResultArray resultArray);
                if (resultArray != null &&
                    resultArray.Size > 0)
                {
                    foreach (IntersectionResult result in resultArray)
                        if (result != null)
                            return result.XYZPoint;
                }
            }
            return null;
        }

        public static void GetInformationConectorWye(FamilyInstance fitting, XYZ vector, out Connector main1, out Connector main2, out Connector tee)
        {
            main1 = null;
            main2 = null;
            tee = null;
            if (fitting != null)
            {
                //Get fitting info

                GetConnectorMain(fitting, vector, out main1, out main2);

                foreach (Connector c in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    if (c.Id != main1.Id && c.Id != main2.Id)
                    {
                        tee = c;
                        break;
                    }
                }
            }
        }

        public static Pipe GetNextPipe(Pipe mainPipe, Pipe splitPipe, XYZ orgPoint)
        {
            double distance1 = mainPipe.ConnectorManager.Lookup(0).Origin.DistanceTo(orgPoint);
            double distance2 = mainPipe.ConnectorManager.Lookup(1).Origin.DistanceTo(orgPoint);
            double distance3 = splitPipe.ConnectorManager.Lookup(0).Origin.DistanceTo(orgPoint);
            double distance4 = splitPipe.ConnectorManager.Lookup(1).Origin.DistanceTo(orgPoint);

            List<double> distances = new List<double>() { distance1, distance2, distance3, distance4 };
            double max = distances.Max(x => x);
            if (max == distance1 || max == distance2)
            {
                return mainPipe;
            }
            else
            {
                return splitPipe;
            }
        }

        public static List<Solid> GetAllSolids(Document doc,
                                              Element elem,
                                              bool getInsGeo,
                                              Autodesk.Revit.DB.View view = null)
        {
            Options options = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = false
            };
            if (view != null)
                options.View = view;

            GeometryElement geoElem = elem.get_Geometry(options);
            List<Solid> solids = new List<Solid>();
            GetSolidFromGeometry(doc, geoElem, getInsGeo, ref solids);
            return solids;
        }

        public static List<PlanarFace> GetPlanarFaces(Element ele, XYZ byDirection = null)
        {
            List<PlanarFace> retVal = new List<PlanarFace>();

            List<Solid> lstSolids = GetAllSolids(ele.Document, ele, true);

            foreach (var solid in lstSolids)
            {
                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (var face in solid.Faces)
                    {
                        if (face != null && face is PlanarFace planar)
                        {
                            if (byDirection == null)
                                retVal.Add(planar);
                            else
                            {
                                if (Common.IsParallel(planar.FaceNormal, byDirection))
                                    retVal.Add(planar);
                            }
                        }
                    }
                }
            }

            return retVal;
        }

        public static XYZ GetCenterElement(Element ele)
        {
            if (ele == null)
                return null;

            // Get bounding box
            var bb = ele.get_BoundingBox(null);
            if (bb == null)
            {
                // Get location
                Location lc = ele.Location;
                if (lc == null)
                    return null;

                if (lc is LocationPoint)
                {
                    // Get location point
                    LocationPoint lcP = lc as LocationPoint;
                    if (lc == null)
                        return null;

                    // Point center
                    var centerP = lcP.Point;
                    return new XYZ(centerP.X, centerP.Y, centerP.Z);
                }
                else if (lc is LocationCurve)
                {
                    // Get location curve
                    LocationCurve lcCurve = lc as LocationCurve;
                    if (lcCurve == null)
                        return null;

                    // Point center
                    var centerP = (lcCurve.Curve.GetEndPoint(1) + lcCurve.Curve.GetEndPoint(0)) / 2;
                    return new XYZ(centerP.X, centerP.Y, centerP.Z);
                }
            }
            else
            {
                XYZ max = bb.Transform.OfPoint(bb.Max);
                XYZ min = bb.Transform.OfPoint(bb.Min);

                // Point center
                var centerP = (max + min) / 2;
                return centerP;
            }

            return null;
        }

        public static bool IsEqual(double first, double second, double tolerance = 1e-5)
        {
            double result = Math.Abs(first - second);
            return result < tolerance;
        }

        public static bool IsOverlap(Line line, PlanarFace plFace)
        {
            if (line == null || plFace == null)
                return false;

            Plane plane = Plane.CreateByNormalAndOrigin(plFace.FaceNormal, plFace.Origin);
            plane.Project(line.Origin, out UV uv, out double distance);
            if (uv != null && IsEqual(distance, 0))
                return true;

            return false;
        }

        public static bool IsSolidGraphicallyVisible(Document doc, Autodesk.Revit.DB.View view, Solid solid)
        {
            if (doc != null
                && view != null
                && solid.GraphicsStyleId != null
                && solid.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (doc.GetElement(solid.GraphicsStyleId) is GraphicsStyle graphicalStyle
                    && graphicalStyle.GraphicsStyleCategory != null)
                    return graphicalStyle.GraphicsStyleCategory.get_Visible(view);
            }
            return true;
        }

        public static void GetSolidFromGeometry(Document doc,
                                                GeometryElement geoElem,
                                                bool getInstGeo,
                                                ref List<Solid> solids,
                                                Autodesk.Revit.DB.View view = null)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Solid solid
                    && solid.Volume > 0
                    && IsSolidGraphicallyVisible(doc, view, solid))
                    solids.Add(solid);
                else if (geoObj is GeometryInstance geoInst)
                {
                    GeometryElement innerGeo = getInstGeo ? geoInst.GetInstanceGeometry() : geoInst.GetSymbolGeometry();
                    GetSolidFromGeometry(doc, innerGeo, getInstGeo, ref solids, view);
                }
            }
        }

        public static List<Curve> GetAllCurves(PlanarFace planarFace)
        {
            List<Curve> retval = new List<Curve>();

            if (planarFace != null)
            {
                foreach (var curveL in planarFace.GetEdgesAsCurveLoops())
                {
                    foreach (var c in curveL)
                    {
                        retval.Add(c);
                    }
                }
            }

            return retval;
        }

        public static List<Curve> GetAllCurves(Solid solid)
        {
            List<Curve> retval = new List<Curve>();

            if (solid != null)
            {
                foreach (Edge edge in solid.Edges)
                {
                    retval.Add(edge.AsCurve());
                }
            }

            return retval;
        }

        public static void GetConnectorMain(FamilyInstance fitting, XYZ vector, out Connector mainConnect1, out Connector mainConnect2)
        {
            mainConnect1 = null;
            mainConnect2 = null;

            if (vector == null && fitting.MEPModel.ConnectorManager.Connectors.Size == 3)
            {
                //Main : hướng connector của 2 connector fai song song voi nhau (nguoc chieu nhau)

                foreach (Connector c1 in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    foreach (Connector c2 in fitting.MEPModel.ConnectorManager.Connectors)
                    {
                        if (c1.Id == c2.Id)
                        {
                            continue;
                        }
                        else
                        {
                            var z1 = c1.CoordinateSystem.BasisZ;
                            var z2 = c2.CoordinateSystem.BasisZ;

                            if (Common.IsParallel(z1, z2, 0.0001) == true)
                            {
                                mainConnect1 = c1;
                                mainConnect2 = c2;
                                break;
                            }
                        }
                    }

                    if (mainConnect1 != null && mainConnect2 != null)
                        break;
                }
            }
            else
            {
                foreach (Connector con in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    if (vector != null)
                    {
                        if (Common.IsParallel(vector, con.CoordinateSystem.BasisZ, 0.0001) == false)
                        {
                            continue;
                        }
                    }

                    if (mainConnect1 == null)
                        mainConnect1 = con;
                    else
                    {
                        mainConnect2 = con;
                        break;
                    }
                }
            }

            if (mainConnect1 != null && mainConnect2 != null)
            {
                //Connect nao gan location of fitting thi do la 1

                var p = (fitting.Location as LocationPoint).Point;
                if (mainConnect1.Origin.DistanceTo(p) > mainConnect2.Origin.DistanceTo(p))
                {
                    Connector temp = mainConnect1;
                    mainConnect1 = mainConnect2;

                    mainConnect2 = temp;
                }
            }
        }

        public static XYZ GetPointProjectOnPlane(Plane plane, XYZ point, XYZ vectorAlong, double Tolerance = 1.0e-5)
        {
            if (vectorAlong != null)
            {
                // Calculate t parameter for intersection
                double numerator = plane.Normal.DotProduct(plane.Origin - point);
                double denominator = plane.Normal.DotProduct(vectorAlong);

                if (Math.Abs(denominator) < Tolerance)
                {
                    return null; // The vector is parallel to the plane
                }

                double t = numerator / denominator;

                // Calculate projected point
                return point + vectorAlong * t;
            }
            else
            {
                plane.Project(point, out UV uv1, out double d);
                XYZ projectedPoint = plane.Origin + (uv1.U * plane.XVec) + (uv1.V * plane.YVec);
                return projectedPoint;
            }
        }

        public static XYZ GetPointProjectOnLine(Line line, XYZ point, bool isMakeUnbound = true)
        {
            if (line != null && point != null)
            {
                Line lineCopy = line.Clone() as Line;

                if (isMakeUnbound)
                    lineCopy.MakeUnbound();

                IntersectionResult intersectionResult = lineCopy.Project(point);
                if (intersectionResult != null)
                {
                    return intersectionResult.XYZPoint;
                }
            }

            return null;
        }

        public static Line GetLineProjectOnPlane(Plane plane, Line line, XYZ vectorAlong = null)
        {
            Line lineProject = null;
            XYZ p1 = GetPointProjectOnPlane(plane, line.GetEndPoint(0), vectorAlong);
            XYZ p2 = GetPointProjectOnPlane(plane, line.GetEndPoint(1), vectorAlong);
            lineProject = Line.CreateBound(p1, p2);
            return lineProject;
        }

        public static XYZ GetPointProjectOnPlane(PlanarFace planar, XYZ point, XYZ vectorAlong, double Tolerance = 1.0e-5)
        {
            Plane plane = Plane.CreateByNormalAndOrigin(planar.FaceNormal, planar.Origin);

            if (vectorAlong != null)
            {
                // Calculate t parameter for intersection
                double numerator = plane.Normal.DotProduct(plane.Origin - point);
                double denominator = plane.Normal.DotProduct(vectorAlong);

                if (Math.Abs(denominator) < Tolerance)
                {
                    return null; // The vector is parallel to the plane
                }

                double t = numerator / denominator;

                // Calculate projected point
                return point + vectorAlong * t;
            }
            else
            {
                plane.Project(point, out UV uv1, out double d);
                XYZ projectedPoint = plane.Origin + (uv1.U * plane.XVec) + (uv1.V * plane.YVec);
                return projectedPoint;
            }
        }

        public static XYZ GetPointIntersecNotInXYPlane(Line line1, Line line2, bool isOnMEPCurveFirst = true)
        {
            if (IsParallel(line1.Direction, line2.Direction))
                return null;

            XYZ normal = line1.Direction.CrossProduct(line2.Direction);

            XYZ origin = line1.Origin;
            if (!isOnMEPCurveFirst)
                origin = line2.Origin;

            Plane plane = Plane.CreateByNormalAndOrigin(normal, origin);

            Line proLine1 = Common.GetLineProjectOnPlane(plane, line1);

            Line proLine2 = Common.GetLineProjectOnPlane(plane, line2);

            proLine1.MakeUnbound();
            proLine2.MakeUnbound();

            IntersectionResultArray iResult = new IntersectionResultArray();
            SetComparisonResult setComparisonResult = proLine1.Intersect(proLine2, out iResult);
            if (setComparisonResult != SetComparisonResult.Disjoint && iResult != null && iResult.Size > 0)
            {
                return iResult.get_Item(0).XYZPoint;
            }

            return null;
        }

        public static List<PlanarFace> GetAllFaceFromElementByDirection(Element ele, XYZ vectorDirection, Transform transform = null)
        {
            if (ele == null)
                return new List<PlanarFace>();

            List<PlanarFace> AllFaces = new List<PlanarFace>();
            foreach (var solid in GetSolidsFromElement(ele, transform))
            {
                foreach (var planarFace in GetSolidFaces(solid))
                    AllFaces.Add(planarFace);
            }

            List<PlanarFace> AllFaceFind = new List<PlanarFace>();

            foreach (var face in AllFaces)
            {
                if (face == null || face.FaceNormal == null)
                    continue;
                XYZ faceNormal = face.FaceNormal.Normalize();

                if (vectorDirection == null)
                    AllFaceFind.Add(face);
                else
                {
                    if (faceNormal.IsAlmostEqualTo(vectorDirection))
                        AllFaceFind.Add(face);
                }
            }

            return AllFaceFind;
        }

        public static List<PlanarFace> GetSolidFaces(Solid solid)
        {
            List<PlanarFace> retVal = new List<PlanarFace>();
            if (solid != null && solid.Faces.Size > 0)
            {
                foreach (var face in solid.Faces)
                {
                    if (face != null)
                        retVal.Add(face as PlanarFace);
                }
            }
            return retVal;
        }

        public static List<Solid> GetSolidsFromElement(Element elem, Transform transform = null)
        {
            var lstSolids = new List<Solid>();
            if (elem == null)
                return new List<Solid>();

            var option = new Options();

            var geoElem = elem.get_Geometry(option);

            if (transform != null)
                geoElem = elem.get_Geometry(option).GetTransformed(transform);

            lstSolids = GetSolidFromGeometryElement(geoElem);
            return lstSolids;
        }

        /// <summary>
        /// Get solid from geometry of element
        /// </summary>
        /// <param name="geomElm"></param>
        /// <returns></returns>
        public static List<Solid> GetSolidFromGeometryElement(GeometryElement geomElm)
        {
            var lstSolids = new List<Solid>();

            if (geomElm == null)
                return new List<Solid>();

            foreach (GeometryObject geoObj in geomElm)
            {
                if (geoObj is GeometryInstance)
                {
                    GeometryInstance geoInst = (GeometryInstance)geoObj;
                    var transInst = geoInst.Transform;
                    var geoElmInst = geoInst.GetInstanceGeometry();
                    lstSolids.AddRange(GetSolidFromGeometryElement(geoElmInst));
                }
                else if (geoObj is Solid)
                {
                    Solid geoSolid = (Solid)geoObj;
                    lstSolids.Add(geoSolid);
                }
            }

            return lstSolids;
        }

        public static bool IsBetweenLine(Line line, XYZ pointCheck)
        {
            if (line != null && line.IsBound == true && pointCheck != null)
            {
                XYZ st = line.GetEndPoint(0);
                XYZ end = line.GetEndPoint(1);

                if (st.IsAlmostEqualTo(pointCheck)
               || end.IsAlmostEqualTo(pointCheck))
                    return false;

                XYZ vec1 = (pointCheck - st).Normalize();

                XYZ vec2 = (pointCheck - end).Normalize();

                if (vec1.DotProduct(vec2) < 0)
                    return true;
            }

            return false;
        }

        public static bool IsPointOnLine(Line line, XYZ pointCheck, double tolerance = 1e-5)
        {
            if (line != null && pointCheck != null)
            {
                XYZ p = GetPointProjectOnLine(line, pointCheck);

                p = p.To2D(pointCheck.Z);

                double distance = p.DistanceTo(pointCheck);
                if (IsEqual(distance, 0, tolerance) || IsEqual(distance, tolerance))
                    return true;
            }

            return false;
        }

        public static bool IsSameDirection(XYZ first, XYZ second, double tolerance = 1e-6)
        {
            double length = first.DotProduct(second);
            return length > 0;
        }

        public static bool IsParallel(XYZ p, XYZ q, double tolerance = 10e-5)
        {
            if (p.CrossProduct(q).IsZeroLength() == true)
                return true;

            var l = p.CrossProduct(q).GetLength();
            if (IsZero(l, tolerance))
                return true;

            return false;
        }

        public static bool IsZero(double a, double tolerance)
        {
            return tolerance > Math.Abs(a);
        }

        public static bool IsEqual(double first, double second)
        {
            double result = Math.Abs(first - second);
            return result < 10e-5;
        }

        /// <summary>
        /// Show information to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowInfor(string message, string title = "情報")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Show warning to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowWarning(string message, string title = "警告")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Show error to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowError(string message, string title = "エラー")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}