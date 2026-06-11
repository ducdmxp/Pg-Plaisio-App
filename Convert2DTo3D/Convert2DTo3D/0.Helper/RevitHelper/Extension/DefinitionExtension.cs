using Autodesk.Revit.DB;

namespace Convert2DTo3D
{
    public static class DefinitionExtension
    {
#if (REVIT2020 || REVIT2021)
        public static ParameterType GetParameterType(this Definition def)
        {
            return def.ParameterType;
        }
#else

        public static ForgeTypeId GetParameterType(this Definition def)
        {
            return def.GetDataType();
        }

#endif
#if (REVIT2020 || REVIT2021)
        public static BuiltInParameterGroup GetParameterGroup(this Definition def)
        {
            return def.ParameterGroup;
        }
#else

        public static ForgeTypeId GetParameterGroup(this Definition def)
        {
            return def.GetGroupTypeId();
        }

#endif
    }
}