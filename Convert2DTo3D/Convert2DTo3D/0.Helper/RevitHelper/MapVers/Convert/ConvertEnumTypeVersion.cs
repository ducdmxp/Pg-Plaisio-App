using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Convert2DTo3D
{
    public class ConvertEnumTypeVersion
    {
#if REVIT2019 || REVIT2020
       public static List<BuiltInCategory> GetAllCategories()
        {
            return Enum.GetValues(typeof(BuiltInCategory)).Cast<BuiltInCategory>().ToList();
        }
        public static List<ParameterType> GetAllParameterTypes()
        {
            return Enum.GetValues(typeof(ParameterType)).Cast<ParameterType>().ToList();
        }
        public static List<BuiltInParameterGroup> GetAllParameterGroup()
        {
            return Enum.GetValues(typeof(BuiltInParameterGroup)).Cast<BuiltInParameterGroup>().ToList();
        }
        public static List<BuiltInParameter> GetAllBuiltInParameter()
        {
            return Enum.GetValues(typeof(BuiltInParameter)).Cast<BuiltInParameter>().ToList();
        }
        public static List<DisplayUnitType> GetAllDisplayUnitType()
        {
            return Enum.GetValues(typeof(DisplayUnitType)).Cast<DisplayUnitType>().ToList();
        }
        public static List<UnitSymbolType> GetAllSymbolTypes()
        {
            return Enum.GetValues(typeof(UnitSymbolType)).Cast<UnitSymbolType>().ToList();
        }
        public static List<UnitType> GetAllUnitTypes()
        {
            return Enum.GetValues(typeof(UnitType)).Cast<UnitType>().ToList();
        }
#elif REVIT2021
        public static List<BuiltInCategory> GetAllCategories()
        {
            return Enum.GetValues(typeof(BuiltInCategory)).Cast<BuiltInCategory>().ToList();
        }
        public static List<ParameterType> GetAllParameterTypes()
        {
            return Enum.GetValues(typeof(ParameterType)).Cast<ParameterType>().ToList();
        }
        public static List<BuiltInParameterGroup> GetAllParameterGroup()
        {
            return Enum.GetValues(typeof(BuiltInParameterGroup)).Cast<BuiltInParameterGroup>().ToList();
        }
        public static List<BuiltInParameter> GetAllBuiltInParameter()
        {
            return Enum.GetValues(typeof(BuiltInParameter)).Cast<BuiltInParameter>().ToList();
        }
        public static List<ForgeTypeId> GetAllDisplayUnitType()
        {
            List<ForgeTypeId> forgeTypeIds = new List<ForgeTypeId>();
            GetAllStaticProperties(typeof(UnitTypeId), ref forgeTypeIds);
            return forgeTypeIds;
        }
        public static List<ForgeTypeId> GetAllSymbolTypes()
        {
            List<ForgeTypeId> forgeTypeIds = new List<ForgeTypeId>();
            GetAllStaticProperties(typeof(SymbolTypeId), ref forgeTypeIds);
            return forgeTypeIds;
        }
        public static List<ForgeTypeId> GetAllUnitTypes()
        {
            return UnitUtils.GetAllSpecs().ToList();
        }

        private static void GetAllStaticProperties(Type myClassType, ref List<ForgeTypeId> result)
        {
            PropertyInfo[] staticFields = myClassType.GetProperties(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            foreach (PropertyInfo field in staticFields)
            {
                object value = field.GetValue(null);
                ForgeTypeId forgeTypeId = value as ForgeTypeId;
                if (forgeTypeId == null) continue;
                result.Add(forgeTypeId);
            }
            Type[] nestedTypes = myClassType.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic);
            foreach (Type nestedType in nestedTypes)
            {
                GetAllStaticProperties(nestedType, ref result);
            }
        }
#else

        public static List<ForgeTypeId> GetAllCategories()
        {
            return Enum.GetValues(typeof(BuiltInCategory)).Cast<BuiltInCategory>().Where(k =>
            {
                try
                {
                    Category.GetBuiltInCategoryTypeId(k);
                    return true;
                }
                catch (Exception)
                { }
                return false;
            }).Select(k =>
                Category.GetBuiltInCategoryTypeId(k)
            ).ToList();
        }

        /// <summary>
        /// Text, Integer, Material, Volumn,....
        /// </summary>
        /// <returns></returns>
        public static List<ForgeTypeId> GetAllParameterTypes()
        {
            return SpecUtils.GetAllSpecs().ToList();
        }

        /// <summary>
        /// PG_Data, PG_Other
        /// </summary>
        /// <returns></returns>
        public static List<ForgeTypeId> GetAllParameterGroup()
        {
            return ParameterUtils.GetAllBuiltInGroups().ToList();
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        public static List<ForgeTypeId> GetAllBuiltInParameter()
        {
            return ParameterUtils.GetAllBuiltInParameters().ToList();
        }

        public static List<ForgeTypeId> GetAllDisplayUnitType()
        {
            List<ForgeTypeId> forgeTypeIds = new List<ForgeTypeId>();
            GetAllStaticProperties(typeof(UnitTypeId), ref forgeTypeIds);
            return forgeTypeIds;
        }

        public static List<ForgeTypeId> GetAllSymbolTypes()
        {
            List<ForgeTypeId> forgeTypeIds = new List<ForgeTypeId>();
            GetAllStaticProperties(typeof(SymbolTypeId), ref forgeTypeIds);
            return forgeTypeIds;
        }

        public static List<ForgeTypeId> GetAllUnitTypes()
        {
            // List<ForgeTypeId> forgeTypeIds = new List<ForgeTypeId>();
            // GetAllStaticProperties(typeof(SpecTypeId), ref forgeTypeIds);
            // return forgeTypeIds;
            // return UnitUtils.GetAllSpecs().ToList();
            return null;
        }

        private static void GetAllStaticProperties(Type myClassType, ref List<ForgeTypeId> result)
        {
            PropertyInfo[] staticFields = myClassType.GetProperties(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            foreach (PropertyInfo field in staticFields)
            {
                object value = field.GetValue(null);
                ForgeTypeId forgeTypeId = value as ForgeTypeId;
                if (forgeTypeId == null) continue;
                result.Add(forgeTypeId);
            }
            Type[] nestedTypes = myClassType.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic);
            foreach (Type nestedType in nestedTypes)
            {
                GetAllStaticProperties(nestedType, ref result);
            }
        }

#endif
    }
}