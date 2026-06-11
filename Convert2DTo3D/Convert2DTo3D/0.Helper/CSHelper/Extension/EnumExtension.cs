using System;

namespace Convert2DTo3D
{
    public static class EnumExtension
    {
        /// <summary>
        /// Extension method to return an enum value of type T for the given string.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T ToEnum<T>(this string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }

        /// <summary>
        /// Extension method to return an enum value of type T for the given int.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        //public static T ToEnum<T>(this int value)
        //{
        //    if (Enum.IsDefined(typeof(T), value))
        //    {
        //        var name = Enum.GetName(typeof(T), value);
        //        return name.ToEnum<T>();
        //    }
        //    else
        //    {
        //        return default(T);
        //    }
        //}
        public static T ToEnum<T>(this int value, out bool res)
        {
            res = true;
            try
            {
                //if (Enum.IsDefined(typeof(T), value))
                //{
                var name = Enum.GetName(typeof(T), value);
                return name.ToEnum<T>();
                //}
                //else
                //{
                //    res = false;
                //    return default(T);
                //}
            }
            catch (Exception ex)
            {
                res = false;
                return default(T);
            }
        }

        public static bool IsDefinedEnum<T>(this string value)
        {
            return Enum.IsDefined(typeof(T), value);
        }
    }
}