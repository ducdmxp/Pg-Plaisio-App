using Autodesk.Revit.DB;
using System.Linq;

namespace Convert2DTo3D
{
    public static class FormatOptionsExtension
    {
#if !(REVIT2019 || REVIT2020)

        public static ForgeTypeId GetDisplayUnitType(this FormatOptions formatOptions)
        {
            return formatOptions.GetUnitTypeId();
        }

        public static ForgeTypeId GetUnitSymbol(this FormatOptions formatOptions)
        {
            return formatOptions.GetValidSymbols().FirstOrDefault();
        }

#else
    public static DisplayUnitType GetDisplayUnitType(this FormatOptions formatOptions)
        {
            return formatOptions.DisplayUnits;
        }
        public static DisplayUnitType GetUnitSymbol(this FormatOptions formatOptions)
        {
            return formatOptions.DisplayUnits;
        }

#endif
    }
}