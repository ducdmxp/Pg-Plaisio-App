using Autodesk.Revit.DB;

namespace Convert2DTo3D
{
    public class LabelVersionUtils
    {
#if REVIT2019 || REVIT2020
        public static string GetLabelFor(BuiltInCategory builtInCategory){
            return LabelUtils.GetLabelFor(builtInCategory);
    }
    public static string GetLabelFor(BuiltInParameter builtInParam){
            return LabelUtils.GetLabelFor(builtInParam);
    }
    public static string GetLabelFor(BuiltInParameterGroup builtInParamGroup)
    {
        return LabelUtils.GetLabelFor(builtInParamGroup);
    }
    public static string GetLabelFor(DisplayUnitType displayUnitType){
       return LabelUtils.GetLabelFor(displayUnitType);
    }
    public static string GetLabelFor(ParameterType paramType){
        return LabelUtils.GetLabelFor(paramType);
     }
     public static string GetLabelFor(UnitSymbolType unitSymbolType){
        return LabelUtils.GetLabelFor(unitSymbolType);
     }
     public static string GetLabelFor(UnitType unitType){
        return LabelUtils.GetLabelFor(unitType);
     }
#elif REVIT2021
        public static string GetLabelFor(ForgeTypeId forgeTypeId)
        {
            try
            {
                if (UnitUtils.IsSymbol(forgeTypeId)) // UnitSymbolType
                {
                    return LabelUtils.GetLabelForSymbol(forgeTypeId);
                }
                else if (UnitUtils.IsUnit(forgeTypeId)) // DisplayUnitType
                {
                    return LabelUtils.GetLabelForUnit(forgeTypeId);
                }
                else  // ParameterType, UnitType
                {
                    return LabelUtils.GetLabelForSpec(forgeTypeId);
                }
            }
            catch (Exception)
            { }
            return "";
        }
        public static string GetLabelFor(ParameterType paramType)
        {
            try
            {
                return LabelUtils.GetLabelFor(paramType);
            }
            catch (Exception)
            {  }
            return "";
        }
        public static string GetLabelFor(BuiltInParameterGroup builtInParamGroup)
        {
            return LabelUtils.GetLabelFor(builtInParamGroup);
        }
        public static string GetLabelFor(BuiltInParameter builtInParam)
        {
            return LabelUtils.GetLabelFor(builtInParam);
        }
        public static string GetLabelFor(BuiltInCategory builtInCategory)
        {
            try
            {
                return LabelUtils.GetLabelFor(builtInCategory);
            }
            catch (Exception)
            {}
            return "";
        }
#else

        public static string GetLabelFor(ForgeTypeId forgeTypeId)
        {
            if (UnitUtils.IsSymbol(forgeTypeId)) // UnitSymbolType
            {
                return LabelUtils.GetLabelForSymbol(forgeTypeId);
            }
            else if (UnitUtils.IsUnit(forgeTypeId)) // DisplayUnitType
            {
                return LabelUtils.GetLabelForUnit(forgeTypeId);
            }
            else if (SpecUtils.IsSpec(forgeTypeId)) // ParameterType, UnitType
            {
                return LabelUtils.GetLabelForSpec(forgeTypeId);
            }
            else if (ParameterUtils.IsBuiltInGroup(forgeTypeId)) // BuiltInParameterGroup
            {
                return LabelUtils.GetLabelForGroup(forgeTypeId);
            }
            else if (ParameterUtils.IsBuiltInParameter(forgeTypeId)) // BuiltInParameter
            {
                return LabelUtils.GetLabelForBuiltInParameter(forgeTypeId);
            }
            else if (Category.IsBuiltInCategory(forgeTypeId)) // BuiltInCategory
            {
                var builtInCategory = Category.GetBuiltInCategory(forgeTypeId);
                return LabelUtils.GetLabelFor(builtInCategory);
            }
            return "";
        }

#endif
    }
}