using System;
using System.Linq;
using System.Reflection;

namespace Convert2DTo3D
{
    public class ObjectUtils
    {
        public static void Copy(object objOrigin, object objCopy, bool isExclude = false, params string[] propertiesName)
        {
            if (objOrigin == null || objCopy == null) return;
            var properties = string.Join("_", propertiesName);
            PropertyInfo[] propertyArr = objOrigin.GetType().GetProperties();
            for (int j = 0; j < propertyArr.Count(); j++)
            {
                string name = propertyArr[j].Name.ToString();
                if (isExclude == true && properties.Contains(name)) continue;
                if ((isExclude == false) && (!(properties.Length > 0 && properties.Contains(name) || properties.Length == 0))) continue;
                if (propertyArr[j].CanWrite)
                {
                    try
                    {
                        objCopy.GetType().GetProperty(name)?.SetValue(objCopy, propertyArr[j].GetValue(objOrigin));
                    }
                    catch (Exception ex)
                    { }
                }
            }
        }

        public static bool SetValue(object objectOrigin, string propertyName, string value)
        {
            PropertyInfo property = objectOrigin.GetType().GetProperty(propertyName);
            if (property != null)
            {
                object convertedValue = Convert.ChangeType(value, property.PropertyType);
                property.SetValue(objectOrigin, convertedValue);
                return true;
            }
            return false;
        }
    }
}