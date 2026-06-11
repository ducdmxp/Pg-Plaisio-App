using Autodesk.Revit.DB;
using System;
using System.Linq;

namespace Convert2DTo3D.Utils
{
    public class ParameterUtils
    {
        /// <summary>
        /// Set giá trị cho paramter theo tên
        /// </summary>
        /// <param name="el"></param>
        /// <param name="parameterName"></param>
        /// <param name="valuePara"></param>
        /// <returns></returns>
        public static bool SetValueParameterByName(Element el, string parameterName, object valuePara)
        {
            if (el == null || string.IsNullOrEmpty(parameterName) || valuePara == null)
                return false;
            Parameter prm = el.LookupParameter(parameterName);
            if (prm != null && !prm.IsReadOnly)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    prm.Set((ElementId)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Double)
                {
                    prm.Set((double)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    prm.Set((int)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.String)
                {
                    prm.Set((string)valuePara);

                    return true;
                }
            }
            return false;
        }

        public static Parameter GetParameterFromListedNames(Element el, params string[] args)
        {
            foreach (string str in args)
            {
                Parameter prm = el.LookupParameter(str);
                if (prm != null)
                {
                    return prm;
                }
            }

            return null;
        }

        /// <summary>
        ///  sét giá trị parameter theo BuiltInParameter
        /// </summary>
        /// <param name="element"></param>
        /// <param name="builtIn"></param>
        /// <param name="valuePara"></param>
        /// <returns></returns>
        public static bool SetValueParameterByBuiltIn(Element element, BuiltInParameter builtIn, object valuePara)
        {
            if (element == null || valuePara == null)
                return false;
            Parameter prm = element.get_Parameter(builtIn);
            if (prm != null && !prm.IsReadOnly)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    prm.Set((ElementId)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Double)
                {
                    prm.Set((double)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    prm.Set((int)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.String)
                {
                    prm.Set((string)valuePara);

                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Lấy giá trị của paramter theo built in parameter
        /// </summary>
        /// <param name="element"></param>
        /// <param name="buintinparameter"></param>
        /// <returns></returns>
        public static object GetValueParameterByBuilt(Element element, BuiltInParameter buintinparameter)
        {
            if (element == null)
                return null;
            Parameter prm = element.get_Parameter(buintinparameter);
            if (prm != null)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    return prm.AsElementId();
                }
                if (prm.StorageType == StorageType.Double)
                {
                    return prm.AsDouble();
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    return prm.AsInteger();
                }
                if (prm.StorageType == StorageType.String)
                {
                    return prm.AsString();
                }
            }
            return null;
        }

        /// <summary>
        /// Set giá trị cho paramter theo tên
        /// </summary>
        /// <param name="el"></param>
        /// <param name="parameterName"></param>
        /// <param name="valuePara"></param>
        /// <returns></returns>
        public static bool SetValueParameterByListName(Element el, object valuePara, params string[] args)
        {
            if (el == null || args == null || args.Count() <= 0 || valuePara == null)
                return false;

            foreach (string str in args)
            {
                Parameter prm = el.LookupParameter(str);
                if (prm != null && !prm.IsReadOnly)
                {
                    if (prm.StorageType == StorageType.ElementId)
                    {
                        prm.Set((ElementId)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.Double)
                    {
                        prm.Set((double)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.Integer)
                    {
                        prm.Set((int)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.String)
                    {
                        prm.Set((string)valuePara);

                        return true;
                    }
                }
            }

            return false;
        }

        public static Object GetValueParameterFromListedNames(Element el, params string[] args)
        {
            foreach (string str in args)
            {
                Object objValue = GetValueParameterByName(el, str);
                if (objValue != null)
                    return objValue;
            }

            return null;
        }

        /// <summary>
        /// Lấy giá trị của paramter theo tên parameter
        /// </summary>
        /// <param name="element"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        public static object GetValueParameterByName(Element element, string parameterName)
        {
            if (element == null || string.IsNullOrEmpty(parameterName))
                return null;
            Parameter prm = element.LookupParameter(parameterName);
            if (prm != null)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    return prm.AsElementId();
                }
                if (prm.StorageType == StorageType.Double)
                {
                    return prm.AsDouble();
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    return prm.AsInteger();
                }
                if (prm.StorageType == StorageType.String)
                {
                    return prm.AsString();
                }
            }
            return null;
        }
    }
}