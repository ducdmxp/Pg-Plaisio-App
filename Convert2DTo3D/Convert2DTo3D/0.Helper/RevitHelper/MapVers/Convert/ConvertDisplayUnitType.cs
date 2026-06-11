using Autodesk.Revit.DB;
using System;

namespace Convert2DTo3D
{
    #region note

    /*
        List<string> forceTypeNames = new List<string>();
        Type unitTypeIdType = typeof(UnitTypeId);
        PropertyInfo[] properties = unitTypeIdType.GetProperties(BindingFlags.Static | BindingFlags.Public);
        foreach (PropertyInfo property in properties) forceTypeNames.Add(property.Name);
        foreach (int i in Enum.GetValues(typeof(DisplayUnitTypeMap)))
        {
            var enumValuen = Enum.GetName(typeof(DisplayUnitTypeMap), i);
            enumValuen = enumValuen.Replace("DUT_", "").ToLower();
            var enumValue= enumValuen.Split('_');

            var name = string.Join("", enumValue) ;
            var chek = forceTypeNames.Where(k => k.ToLower().Contains(name) || k.ToLower() == name).ToList();

            if (chek.Count() == 0)
            {
                Debug.WriteLine($"case BuiltInParameterMap.{Enum.GetName(typeof(DisplayUnitTypeMap), i)}: \n\treturn UnitTypeId.Valid;");
            }
            else
            {
                string ne = "";
                foreach (var item in chek)
                {
                    ne += item;
                    if (chek.Count() > 1)
                    {
                        ne += "|";
                    }
                }
                ne.Remove(ne.Count() - 1, 1);
                Debug.WriteLine($"case BuiltInParameterMap.{Enum.GetName(typeof(DisplayUnitTypeMap), i)}: \n\t  return UnitTypeId.{ne};");
            }
        }
     */

    #endregion note

    public class ConvertDisplayUnitType
    {
#if (REVIT2018 || REVIT2019 || REVIT2020)

        public static DisplayUnitType Convert(DisplayUnitTypeMap displayUnitTypeMap)
        {
            var stringEnum = displayUnitTypeMap.ToString();
            var checkExists = Enum.IsDefined(typeof(DisplayUnitType), stringEnum);
            if (checkExists == true)
            {
                DisplayUnitType output = DisplayUnitType.DUT_MILLIMETERS;
                Enum.TryParse<DisplayUnitType>(stringEnum, out output);
                return output;
            }
            return DisplayUnitType.DUT_MILLIMETERS;
        }

#else

        public static ForgeTypeId Convert(DisplayUnitTypeMap mapParameterType)
        {
            switch (mapParameterType)
            {
                case DisplayUnitTypeMap.DUT_CUSTOM:
                    return UnitTypeId.Custom;

                case DisplayUnitTypeMap.DUT_METERS:
                    return UnitTypeId.Meters; //
                case DisplayUnitTypeMap.DUT_CENTIMETERS:
                    return UnitTypeId.Centimeters;

                case DisplayUnitTypeMap.DUT_MILLIMETERS:
                    return UnitTypeId.Millimeters;

                case DisplayUnitTypeMap.DUT_FEET_FRACTIONAL_INCHES:
                    return UnitTypeId.FeetFractionalInches;

                case DisplayUnitTypeMap.DUT_FRACTIONAL_INCHES:
                    return UnitTypeId.FractionalInches;

                case DisplayUnitTypeMap.DUT_ACRES:
                    return UnitTypeId.Acres;

                case DisplayUnitTypeMap.DUT_HECTARES:
                    return UnitTypeId.Hectares;

                case DisplayUnitTypeMap.DUT_METERS_CENTIMETERS:
                    return UnitTypeId.MetersCentimeters;

                case DisplayUnitTypeMap.DUT_CUBIC_YARDS:
                    return UnitTypeId.CubicYards;

                case DisplayUnitTypeMap.DUT_SQUARE_FEET:
                    return UnitTypeId.SquareFeet;

                case DisplayUnitTypeMap.DUT_SQUARE_METERS:
                    return UnitTypeId.SquareMeters;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET:
                    return UnitTypeId.CubicFeet;

                case DisplayUnitTypeMap.DUT_CUBIC_METERS:
                    return UnitTypeId.CubicMeters;

                case DisplayUnitTypeMap.DUT_DEGREES_AND_MINUTES:
                    return UnitTypeId.DegreesMinutes;

                case DisplayUnitTypeMap.DUT_GENERAL:
                    return UnitTypeId.General;

                case DisplayUnitTypeMap.DUT_FIXED:
                    return UnitTypeId.Fixed;

                case DisplayUnitTypeMap.DUT_PERCENTAGE:
                    return UnitTypeId.Percentage;

                case DisplayUnitTypeMap.DUT_SQUARE_INCHES:
                    return UnitTypeId.SquareInches;

                case DisplayUnitTypeMap.DUT_SQUARE_CENTIMETERS:
                    return UnitTypeId.SquareCentimeters;

                case DisplayUnitTypeMap.DUT_SQUARE_MILLIMETERS:
                    return UnitTypeId.SquareMillimeters;

                case DisplayUnitTypeMap.DUT_CUBIC_INCHES:
                    return UnitTypeId.CubicInches;

                case DisplayUnitTypeMap.DUT_CUBIC_CENTIMETERS:
                    return UnitTypeId.CubicCentimeters;

                case DisplayUnitTypeMap.DUT_CUBIC_MILLIMETERS:
                    return UnitTypeId.CubicMillimeters;

                case DisplayUnitTypeMap.DUT_LITERS:
                    return UnitTypeId.Liters;

                case DisplayUnitTypeMap.DUT_GALLONS_US:
                    return UnitTypeId.UsGallons;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_PER_CUBIC_METER:
                    return UnitTypeId.KilogramsPerCubicMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_CUBIC_FOOT:
                    return UnitTypeId.PoundsMassPerCubicFoot;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_CUBIC_INCH:
                    return UnitTypeId.PoundsMassPerCubicInch;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS:
                    return UnitTypeId.BritishThermalUnits;

                case DisplayUnitTypeMap.DUT_CALORIES:
                    return UnitTypeId.Calories;

                case DisplayUnitTypeMap.DUT_KILOCALORIES:
                    return UnitTypeId.Kilocalories;

                case DisplayUnitTypeMap.DUT_JOULES:
                    return UnitTypeId.Joules;

                case DisplayUnitTypeMap.DUT_KILOWATT_HOURS:
                    return UnitTypeId.KilowattHours;

                case DisplayUnitTypeMap.DUT_THERMS:
                    return UnitTypeId.Therms;

                case DisplayUnitTypeMap.DUT_INCHES_OF_WATER_PER_100FT:
                    return UnitTypeId.InchesOfWater60DegreesFahrenheitPer100Feet;

                case DisplayUnitTypeMap.DUT_PASCALS_PER_METER:
                    return UnitTypeId.PascalsPerMeter;

                case DisplayUnitTypeMap.DUT_WATTS:
                    return UnitTypeId.Watts;

                case DisplayUnitTypeMap.DUT_KILOWATTS:
                    return UnitTypeId.Kilowatts;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_SECOND:
                    return UnitTypeId.BritishThermalUnitsPerSecond;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_HOUR:
                    return UnitTypeId.BritishThermalUnitsPerHour;

                case DisplayUnitTypeMap.DUT_CALORIES_PER_SECOND:
                    return UnitTypeId.CaloriesPerSecond;

                case DisplayUnitTypeMap.DUT_KILOCALORIES_PER_SECOND:
                    return UnitTypeId.KilocaloriesPerSecond;

                case DisplayUnitTypeMap.DUT_WATTS_PER_SQUARE_FOOT:
                    return UnitTypeId.WattsPerSquareFoot;

                case DisplayUnitTypeMap.DUT_WATTS_PER_SQUARE_METER:
                    return UnitTypeId.WattsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_INCHES_OF_WATER:
                    return UnitTypeId.InchesOfWater60DegreesFahrenheit;

                case DisplayUnitTypeMap.DUT_PASCALS:
                    return UnitTypeId.Pascals;

                case DisplayUnitTypeMap.DUT_KILOPASCALS:
                    return UnitTypeId.Kilopascals;

                case DisplayUnitTypeMap.DUT_MEGAPASCALS:
                    return UnitTypeId.Megapascals;

                case DisplayUnitTypeMap.DUT_POUNDS_FORCE_PER_SQUARE_INCH:
                    return UnitTypeId.PoundsForcePerSquareInch;

                case DisplayUnitTypeMap.DUT_INCHES_OF_MERCURY:
                    return UnitTypeId.InchesOfMercury32DegreesFahrenheit;

                case DisplayUnitTypeMap.DUT_MILLIMETERS_OF_MERCURY:
                    return UnitTypeId.MillimetersOfMercury;

                case DisplayUnitTypeMap.DUT_ATMOSPHERES:
                    return UnitTypeId.Atmospheres;

                case DisplayUnitTypeMap.DUT_BARS:
                    return UnitTypeId.Bars;

                case DisplayUnitTypeMap.DUT_FAHRENHEIT:
                    return UnitTypeId.Fahrenheit;

                case DisplayUnitTypeMap.DUT_CELSIUS:
                    return UnitTypeId.Celsius;

                case DisplayUnitTypeMap.DUT_KELVIN:
                    return UnitTypeId.Kelvin;

                case DisplayUnitTypeMap.DUT_RANKINE:
                    return UnitTypeId.Rankine;

                case DisplayUnitTypeMap.DUT_FEET_PER_MINUTE:
                    return UnitTypeId.FeetPerMinute;

                case DisplayUnitTypeMap.DUT_METERS_PER_SECOND:
                    return UnitTypeId.MetersPerSecond;

                case DisplayUnitTypeMap.DUT_CENTIMETERS_PER_MINUTE:
                    return UnitTypeId.CentimetersPerMinute;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET_PER_MINUTE:
                    return UnitTypeId.CubicFeetPerMinute;

                case DisplayUnitTypeMap.DUT_LITERS_PER_SECOND:
                    return UnitTypeId.LitersPerSecond;

                case DisplayUnitTypeMap.DUT_CUBIC_METERS_PER_SECOND:
                    return UnitTypeId.CubicMetersPerSecond;

                case DisplayUnitTypeMap.DUT_CUBIC_METERS_PER_HOUR:
                    return UnitTypeId.CubicMetersPerHour;

                case DisplayUnitTypeMap.DUT_GALLONS_US_PER_MINUTE:
                    return UnitTypeId.UsGallonsPerMinute;

                case DisplayUnitTypeMap.DUT_GALLONS_US_PER_HOUR:
                    return UnitTypeId.UsGallonsPerHour;

                case DisplayUnitTypeMap.DUT_AMPERES:
                    return UnitTypeId.Amperes;

                case DisplayUnitTypeMap.DUT_KILOAMPERES:
                    return UnitTypeId.Kiloamperes;

                case DisplayUnitTypeMap.DUT_MILLIAMPERES:
                    return UnitTypeId.Milliamperes;

                case DisplayUnitTypeMap.DUT_VOLTS:
                    return UnitTypeId.Volts;

                case DisplayUnitTypeMap.DUT_KILOVOLTS:
                    return UnitTypeId.Kilovolts;

                case DisplayUnitTypeMap.DUT_MILLIVOLTS:
                    return UnitTypeId.Millivolts;

                case DisplayUnitTypeMap.DUT_HERTZ:
                    return UnitTypeId.Hertz;

                case DisplayUnitTypeMap.DUT_CYCLES_PER_SECOND:
                    return UnitTypeId.CyclesPerSecond;

                case DisplayUnitTypeMap.DUT_LUX:
                    return UnitTypeId.Lux;

                case DisplayUnitTypeMap.DUT_FOOTCANDLES:
                    return UnitTypeId.Footcandles;

                case DisplayUnitTypeMap.DUT_FOOTLAMBERTS:
                    return UnitTypeId.Footlamberts;

                case DisplayUnitTypeMap.DUT_CANDELAS_PER_SQUARE_METER:
                    return UnitTypeId.CandelasPerSquareMeter;

                case DisplayUnitTypeMap.DUT_CANDELAS:
                    return UnitTypeId.Candelas;

                case DisplayUnitTypeMap.DUT_LUMENS:
                    return UnitTypeId.Lumens;

                case DisplayUnitTypeMap.DUT_VOLT_AMPERES:
                    return UnitTypeId.VoltAmperes;

                case DisplayUnitTypeMap.DUT_KILOVOLT_AMPERES:
                    return UnitTypeId.KilovoltAmperes;

                case DisplayUnitTypeMap.DUT_HORSEPOWER:
                    return UnitTypeId.Horsepower;

                case DisplayUnitTypeMap.DUT_NEWTONS:
                    return UnitTypeId.Newtons;

                case DisplayUnitTypeMap.DUT_DECANEWTONS:
                    return UnitTypeId.Dekanewtons;

                case DisplayUnitTypeMap.DUT_KILONEWTONS:
                    return UnitTypeId.Kilonewtons;

                case DisplayUnitTypeMap.DUT_MEGANEWTONS:
                    return UnitTypeId.Meganewtons;

                case DisplayUnitTypeMap.DUT_KIPS:
                    return UnitTypeId.Kips;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_FORCE:
                    return UnitTypeId.KilogramsForce;

                case DisplayUnitTypeMap.DUT_TONNES_FORCE:
                    return UnitTypeId.TonnesForce;

                case DisplayUnitTypeMap.DUT_POUNDS_FORCE:
                    return UnitTypeId.PoundsForce;

                case DisplayUnitTypeMap.DUT_NEWTONS_PER_METER:
                    return UnitTypeId.NewtonsPerMeter;

                case DisplayUnitTypeMap.DUT_DECANEWTONS_PER_METER:
                    return UnitTypeId.DekanewtonsPerMeter;

                case DisplayUnitTypeMap.DUT_KILONEWTONS_PER_METER:
                    return UnitTypeId.KilonewtonsPerMeter;

                case DisplayUnitTypeMap.DUT_MEGANEWTONS_PER_METER:
                    return UnitTypeId.MeganewtonsPerMeter;

                case DisplayUnitTypeMap.DUT_KIPS_PER_FOOT:
                    return UnitTypeId.KipsPerFoot;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_FORCE_PER_METER:
                    return UnitTypeId.KilogramsForcePerMeter;

                case DisplayUnitTypeMap.DUT_TONNES_FORCE_PER_METER:
                    return UnitTypeId.TonnesForcePerMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_FORCE_PER_FOOT:
                    return UnitTypeId.PoundsForcePerFoot;

                case DisplayUnitTypeMap.DUT_NEWTONS_PER_SQUARE_METER:
                    return UnitTypeId.NewtonsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_DECANEWTONS_PER_SQUARE_METER:
                    return UnitTypeId.DekanewtonsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_KILONEWTONS_PER_SQUARE_METER:
                    return UnitTypeId.KilonewtonsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_MEGANEWTONS_PER_SQUARE_METER:
                    return UnitTypeId.MeganewtonsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_KIPS_PER_SQUARE_FOOT:
                    return UnitTypeId.KipsPerSquareFoot;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_FORCE_PER_SQUARE_METER:
                    return UnitTypeId.KilogramsForcePerSquareMeter;

                case DisplayUnitTypeMap.DUT_TONNES_FORCE_PER_SQUARE_METER:
                    return UnitTypeId.TonnesForcePerSquareMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_FORCE_PER_SQUARE_FOOT:
                    return UnitTypeId.PoundsForcePerSquareFoot;

                case DisplayUnitTypeMap.DUT_NEWTON_METERS:
                    return UnitTypeId.NewtonMeters;

                case DisplayUnitTypeMap.DUT_DECANEWTON_METERS:
                    return UnitTypeId.DekanewtonMeters;

                case DisplayUnitTypeMap.DUT_KILONEWTON_METERS:
                    return UnitTypeId.KilonewtonMeters;

                case DisplayUnitTypeMap.DUT_MEGANEWTON_METERS:
                    return UnitTypeId.MeganewtonMeters;

                case DisplayUnitTypeMap.DUT_KIP_FEET:
                    return UnitTypeId.KipFeet;

                case DisplayUnitTypeMap.DUT_KILOGRAM_FORCE_METERS:
                    return UnitTypeId.KilogramForceMeters;

                case DisplayUnitTypeMap.DUT_TONNE_FORCE_METERS:
                    return UnitTypeId.TonneForceMeters;

                case DisplayUnitTypeMap.DUT_POUND_FORCE_FEET:
                    return UnitTypeId.PoundForceFeet;

                case DisplayUnitTypeMap.DUT_METERS_PER_KILONEWTON:
                    return UnitTypeId.MetersPerKilonewton;

                case DisplayUnitTypeMap.DUT_FEET_PER_KIP:
                    return UnitTypeId.FeetPerKip;

                case DisplayUnitTypeMap.DUT_SQUARE_METERS_PER_KILONEWTON:
                    return UnitTypeId.SquareMetersPerKilonewton;

                case DisplayUnitTypeMap.DUT_SQUARE_FEET_PER_KIP:
                    return UnitTypeId.SquareFeetPerKip;

                case DisplayUnitTypeMap.DUT_CUBIC_METERS_PER_KILONEWTON:
                    return UnitTypeId.CubicMetersPerKilonewton;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET_PER_KIP:
                    return UnitTypeId.CubicFeetPerKip;

                case DisplayUnitTypeMap.DUT_INV_KILONEWTONS:
                    return UnitTypeId.InverseKilonewtons;

                case DisplayUnitTypeMap.DUT_INV_KIPS:
                    return UnitTypeId.InverseKips;

                case DisplayUnitTypeMap.DUT_FEET_OF_WATER_PER_100FT:
                    return UnitTypeId.FeetOfWater39_2DegreesFahrenheitPer100Feet;

                case DisplayUnitTypeMap.DUT_FEET_OF_WATER:
                    return UnitTypeId.FeetOfWater39_2DegreesFahrenheit;

                case DisplayUnitTypeMap.DUT_PASCAL_SECONDS:
                    return UnitTypeId.PascalSeconds;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_FOOT_SECOND:
                    return UnitTypeId.PoundsMassPerFootSecond;

                case DisplayUnitTypeMap.DUT_CENTIPOISES:
                    return UnitTypeId.Centipoises;

                case DisplayUnitTypeMap.DUT_FEET_PER_SECOND:
                    return UnitTypeId.FeetPerSecond;

                case DisplayUnitTypeMap.DUT_KIPS_PER_SQUARE_INCH:
                    return UnitTypeId.KipsPerSquareInch;

                case DisplayUnitTypeMap.DUT_KILONEWTONS_PER_CUBIC_METER:
                    return UnitTypeId.KilonewtonsPerCubicMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_FORCE_PER_CUBIC_FOOT:
                    return UnitTypeId.PoundsForcePerCubicFoot;

                case DisplayUnitTypeMap.DUT_KIPS_PER_CUBIC_INCH:
                    return UnitTypeId.KipsPerCubicInch;

                case DisplayUnitTypeMap.DUT_INV_FAHRENHEIT:
                    return UnitTypeId.InverseDegreesFahrenheit;

                case DisplayUnitTypeMap.DUT_INV_CELSIUS:
                    return UnitTypeId.InverseDegreesCelsius;

                case DisplayUnitTypeMap.DUT_NEWTON_METERS_PER_METER:
                    return UnitTypeId.NewtonMetersPerMeter;

                case DisplayUnitTypeMap.DUT_DECANEWTON_METERS_PER_METER:
                    return UnitTypeId.DekanewtonMetersPerMeter;

                case DisplayUnitTypeMap.DUT_KILONEWTON_METERS_PER_METER:
                    return UnitTypeId.KilonewtonMetersPerMeter;

                case DisplayUnitTypeMap.DUT_MEGANEWTON_METERS_PER_METER:
                    return UnitTypeId.MeganewtonMetersPerMeter;

                case DisplayUnitTypeMap.DUT_KIP_FEET_PER_FOOT:
                    return UnitTypeId.KipFeetPerFoot;

                case DisplayUnitTypeMap.DUT_KILOGRAM_FORCE_METERS_PER_METER:
                    return UnitTypeId.KilogramForceMetersPerMeter;

                case DisplayUnitTypeMap.DUT_TONNE_FORCE_METERS_PER_METER:
                    return UnitTypeId.TonneForceMetersPerMeter;

                case DisplayUnitTypeMap.DUT_POUND_FORCE_FEET_PER_FOOT:
                    return UnitTypeId.PoundForceFeetPerFoot;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_FOOT_HOUR:
                    return UnitTypeId.PoundsMassPerFootHour;

                case DisplayUnitTypeMap.DUT_KIPS_PER_INCH:
                    return UnitTypeId.KipsPerInch;

                case DisplayUnitTypeMap.DUT_KIPS_PER_CUBIC_FOOT:
                    return UnitTypeId.KipsPerCubicFoot;

                case DisplayUnitTypeMap.DUT_KIP_FEET_PER_DEGREE:
                    return UnitTypeId.KipFeetPerDegree;

                case DisplayUnitTypeMap.DUT_KILONEWTON_METERS_PER_DEGREE:
                    return UnitTypeId.KilonewtonMetersPerDegree;

                case DisplayUnitTypeMap.DUT_KIP_FEET_PER_DEGREE_PER_FOOT:
                    return UnitTypeId.KipFeetPerDegreePerFoot;

                case DisplayUnitTypeMap.DUT_KILONEWTON_METERS_PER_DEGREE_PER_METER:
                    return UnitTypeId.KilonewtonMetersPerDegreePerMeter;

                case DisplayUnitTypeMap.DUT_WATTS_PER_SQUARE_METER_KELVIN:
                    return UnitTypeId.WattsPerSquareMeterKelvin;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT_FAHRENHEIT:
                    return UnitTypeId.BritishThermalUnitsPerHourSquareFootDegreeFahrenheit;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET_PER_MINUTE_SQUARE_FOOT:
                    return UnitTypeId.CubicFeetPerMinuteSquareFoot;

                case DisplayUnitTypeMap.DUT_LITERS_PER_SECOND_SQUARE_METER:
                    return UnitTypeId.LitersPerSecondSquareMeter;

                case DisplayUnitTypeMap.DUT_RATIO_10:
                    return UnitTypeId.RatioTo10;

                case DisplayUnitTypeMap.DUT_RATIO_12:
                    return UnitTypeId.RatioTo12;

                case DisplayUnitTypeMap.DUT_SLOPE_DEGREES:
                    return UnitTypeId.SlopeDegrees;

                case DisplayUnitTypeMap.DUT_RISE_OVER_INCHES:
                    return UnitTypeId.RiseDividedBy12Inches; //
                case DisplayUnitTypeMap.DUT_RISE_OVER_FOOT:
                    return UnitTypeId.RiseDividedBy1Foot; //
                case DisplayUnitTypeMap.DUT_RISE_OVER_MMS:
                    return UnitTypeId.RiseDividedBy120Inches; //
                case DisplayUnitTypeMap.DUT_WATTS_PER_CUBIC_FOOT:
                    return UnitTypeId.WattsPerCubicFoot;

                case DisplayUnitTypeMap.DUT_WATTS_PER_CUBIC_METER:
                    return UnitTypeId.WattsPerCubicMeter;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT:
                    return UnitTypeId.BritishThermalUnitsPerHourSquareFoot;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_HOUR_CUBIC_FOOT:
                    return UnitTypeId.BritishThermalUnitsPerHourCubicFoot;

                case DisplayUnitTypeMap.DUT_TON_OF_REFRIGERATION:
                    return UnitTypeId.TonsOfRefrigeration;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET_PER_MINUTE_CUBIC_FOOT:
                    return UnitTypeId.CubicFeetPerMinuteCubicFoot;

                case DisplayUnitTypeMap.DUT_LITERS_PER_SECOND_CUBIC_METER:
                    return UnitTypeId.LitersPerSecondCubicMeter;

                case DisplayUnitTypeMap.DUT_CUBIC_FEET_PER_MINUTE_TON_OF_REFRIGERATION:
                    return UnitTypeId.CubicFeetPerMinuteTonOfRefrigeration;

                case DisplayUnitTypeMap.DUT_LITERS_PER_SECOND_KILOWATTS:
                    return UnitTypeId.LitersPerSecondKilowatt;

                case DisplayUnitTypeMap.DUT_SQUARE_FEET_PER_TON_OF_REFRIGERATION:
                    return UnitTypeId.SquareFeetPerTonOfRefrigeration;

                case DisplayUnitTypeMap.DUT_SQUARE_METERS_PER_KILOWATTS:
                    return UnitTypeId.SquareMetersPerKilowatt;

                case DisplayUnitTypeMap.DUT_CURRENCY:
                    return UnitTypeId.Currency;

                case DisplayUnitTypeMap.DUT_LUMENS_PER_WATT:
                    return UnitTypeId.LumensPerWatt;

                case DisplayUnitTypeMap.DUT_SQUARE_FEET_PER_THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR:
                    return UnitTypeId.SquareFeetPer1000BritishThermalUnitsPerHour;

                case DisplayUnitTypeMap.DUT_KILONEWTONS_PER_SQUARE_CENTIMETER:
                    return UnitTypeId.KilonewtonsPerSquareCentimeter;

                case DisplayUnitTypeMap.DUT_NEWTONS_PER_SQUARE_MILLIMETER:
                    return UnitTypeId.NewtonsPerSquareMillimeter;

                case DisplayUnitTypeMap.DUT_KILONEWTONS_PER_SQUARE_MILLIMETER:
                    return UnitTypeId.KilonewtonsPerSquareMillimeter;

                case DisplayUnitTypeMap.DUT_RISE_OVER_120_INCHES:
                    return UnitTypeId.RiseDividedBy120Inches;

                case DisplayUnitTypeMap.DUT_1_RATIO:
                    return UnitTypeId.OneToRatio;

                case DisplayUnitTypeMap.DUT_RISE_OVER_10_FEET:
                    return UnitTypeId.RiseDividedBy10Feet;

                case DisplayUnitTypeMap.DUT_HOUR_SQUARE_FOOT_FAHRENHEIT_PER_BRITISH_THERMAL_UNIT:
                    return UnitTypeId.HourSquareFootDegreesFahrenheitPerBritishThermalUnit;

                case DisplayUnitTypeMap.DUT_SQUARE_METER_KELVIN_PER_WATT:
                    return UnitTypeId.SquareMeterKelvinsPerWatt;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNIT_PER_FAHRENHEIT:
                    return UnitTypeId.BritishThermalUnitsPerDegreeFahrenheit;

                case DisplayUnitTypeMap.DUT_JOULES_PER_KELVIN:
                    return UnitTypeId.JoulesPerKelvin;

                case DisplayUnitTypeMap.DUT_KILOJOULES_PER_KELVIN:
                    return UnitTypeId.KilojoulesPerKelvin;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_MASS:
                    return UnitTypeId.Kilograms;

                case DisplayUnitTypeMap.DUT_TONNES_MASS:
                    return UnitTypeId.UsTonnesMass;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS:
                    return UnitTypeId.PoundsMass;

                case DisplayUnitTypeMap.DUT_METERS_PER_SECOND_SQUARED:
                    return UnitTypeId.MetersPerSecondSquared;

                case DisplayUnitTypeMap.DUT_KILOMETERS_PER_SECOND_SQUARED:
                    return UnitTypeId.KilometersPerSecondSquared;

                case DisplayUnitTypeMap.DUT_INCHES_PER_SECOND_SQUARED:
                    return UnitTypeId.InchesPerSecondSquared;

                case DisplayUnitTypeMap.DUT_FEET_PER_SECOND_SQUARED:
                    return UnitTypeId.FeetPerSecondSquared;

                case DisplayUnitTypeMap.DUT_MILES_PER_SECOND_SQUARED:
                    return UnitTypeId.MilesPerSecondSquared;

                case DisplayUnitTypeMap.DUT_FEET_TO_THE_FOURTH_POWER:
                    return UnitTypeId.FeetToTheFourthPower;

                case DisplayUnitTypeMap.DUT_INCHES_TO_THE_FOURTH_POWER:
                    return UnitTypeId.InchesToTheFourthPower;

                case DisplayUnitTypeMap.DUT_MILLIMETERS_TO_THE_FOURTH_POWER:
                    return UnitTypeId.MillimetersToTheFourthPower;

                case DisplayUnitTypeMap.DUT_CENTIMETERS_TO_THE_FOURTH_POWER:
                    return UnitTypeId.CentimetersToTheFourthPower;

                case DisplayUnitTypeMap.DUT_METERS_TO_THE_FOURTH_POWER:
                    return UnitTypeId.MetersToTheFourthPower;

                case DisplayUnitTypeMap.DUT_FEET_TO_THE_SIXTH_POWER:
                    return UnitTypeId.FeetToTheSixthPower;

                case DisplayUnitTypeMap.DUT_INCHES_TO_THE_SIXTH_POWER:
                    return UnitTypeId.InchesToTheSixthPower;

                case DisplayUnitTypeMap.DUT_MILLIMETERS_TO_THE_SIXTH_POWER:
                    return UnitTypeId.MillimetersToTheSixthPower;

                case DisplayUnitTypeMap.DUT_CENTIMETERS_TO_THE_SIXTH_POWER:
                    return UnitTypeId.CentimetersToTheSixthPower;

                case DisplayUnitTypeMap.DUT_METERS_TO_THE_SIXTH_POWER:
                    return UnitTypeId.MetersToTheSixthPower;

                case DisplayUnitTypeMap.DUT_SQUARE_FEET_PER_FOOT:
                    return UnitTypeId.SquareFeetPerFoot;

                case DisplayUnitTypeMap.DUT_SQUARE_INCHES_PER_FOOT:
                    return UnitTypeId.SquareInchesPerFoot;

                case DisplayUnitTypeMap.DUT_SQUARE_MILLIMETERS_PER_METER:
                    return UnitTypeId.SquareMillimetersPerMeter;

                case DisplayUnitTypeMap.DUT_SQUARE_CENTIMETERS_PER_METER:
                    return UnitTypeId.SquareCentimetersPerMeter;

                case DisplayUnitTypeMap.DUT_SQUARE_METERS_PER_METER:
                    return UnitTypeId.SquareMetersPerMeter;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_MASS_PER_METER:
                    return UnitTypeId.KilogramsPerMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_FOOT:
                    return UnitTypeId.PoundsMassPerFoot;

                case DisplayUnitTypeMap.DUT_RADIANS:
                    return UnitTypeId.Radians;

                case DisplayUnitTypeMap.DUT_GRADS:
                    return UnitTypeId.Gradians;

                case DisplayUnitTypeMap.DUT_RADIANS_PER_SECOND:
                    return UnitTypeId.RadiansPerSecond;

                case DisplayUnitTypeMap.DUT_MILISECONDS:
                    return UnitTypeId.Milliseconds;

                case DisplayUnitTypeMap.DUT_SECONDS:
                    return UnitTypeId.Seconds;

                case DisplayUnitTypeMap.DUT_MINUTES:
                    return UnitTypeId.Minutes;

                case DisplayUnitTypeMap.DUT_HOURS:
                    return UnitTypeId.Hours;

                case DisplayUnitTypeMap.DUT_KILOMETERS_PER_HOUR:
                    return UnitTypeId.KilometersPerHour;

                case DisplayUnitTypeMap.DUT_MILES_PER_HOUR:
                    return UnitTypeId.MilesPerHour;

                case DisplayUnitTypeMap.DUT_KILOJOULES:
                    return UnitTypeId.Kilojoules;

                case DisplayUnitTypeMap.DUT_KILOGRAMS_MASS_PER_SQUARE_METER:
                    return UnitTypeId.KilogramsPerSquareMeter;

                case DisplayUnitTypeMap.DUT_POUNDS_MASS_PER_SQUARE_FOOT:
                    return UnitTypeId.PoundsMassPerSquareFoot;

                case DisplayUnitTypeMap.DUT_WATTS_PER_METER_KELVIN:
                    return UnitTypeId.WattsPerMeterKelvin;

                case DisplayUnitTypeMap.DUT_JOULES_PER_GRAM_CELSIUS:
                    return UnitTypeId.JoulesPerGramDegreeCelsius;

                case DisplayUnitTypeMap.DUT_JOULES_PER_GRAM:
                    return UnitTypeId.JoulesPerGram;

                case DisplayUnitTypeMap.DUT_NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER:
                    return UnitTypeId.NanogramsPerPascalSecondSquareMeter;

                case DisplayUnitTypeMap.DUT_OHM_METERS:
                    return UnitTypeId.OhmMeters;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_HOUR_FOOT_FAHRENHEIT:
                    return UnitTypeId.BritishThermalUnitsPerHourFootDegreeFahrenheit;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_POUND_FAHRENHEIT:
                    return UnitTypeId.BritishThermalUnitsPerPoundDegreeFahrenheit;

                case DisplayUnitTypeMap.DUT_BRITISH_THERMAL_UNITS_PER_POUND:
                    return UnitTypeId.BritishThermalUnitsPerPound;

                case DisplayUnitTypeMap.DUT_GRAINS_PER_HOUR_SQUARE_FOOT_INCH_MERCURY:
                    return UnitTypeId.GrainsPerHourSquareFootInchMercury;

                case DisplayUnitTypeMap.DUT_PER_MILLE:
                    return UnitTypeId.PerMille;

                case DisplayUnitTypeMap.DUT_DECIMETERS:
                    return UnitTypeId.Decimeters;

                case DisplayUnitTypeMap.DUT_JOULES_PER_KILOGRAM_CELSIUS:
                    return UnitTypeId.JoulesPerKilogramDegreeCelsius;

                case DisplayUnitTypeMap.DUT_MICROMETERS_PER_METER_CELSIUS:
                    return UnitTypeId.MicrometersPerMeterDegreeCelsius;

                case DisplayUnitTypeMap.DUT_MICROINCHES_PER_INCH_FAHRENHEIT:
                    return UnitTypeId.MicroinchesPerInchDegreeFahrenheit;

                case DisplayUnitTypeMap.DUT_USTONNES_MASS:
                    return UnitTypeId.UsTonnesMass;

                case DisplayUnitTypeMap.DUT_USTONNES_FORCE:
                    return UnitTypeId.UsTonnesForce;

                case DisplayUnitTypeMap.DUT_LITERS_PER_MINUTE:
                    return UnitTypeId.LitersPerMinute;

                case DisplayUnitTypeMap.DUT_FAHRENHEIT_DIFFERENCE:
                    return UnitTypeId.FahrenheitInterval; //
                case DisplayUnitTypeMap.DUT_CELSIUS_DIFFERENCE:
                    return UnitTypeId.CelsiusInterval;   //
                case DisplayUnitTypeMap.DUT_KELVIN_DIFFERENCE:
                    return UnitTypeId.KelvinInterval;    //
                case DisplayUnitTypeMap.DUT_RANKINE_DIFFERENCE:
                    return UnitTypeId.RankineInterval;   //
                default:
                    return UnitTypeId.Millimeters;
            }
        }

#endif
    }
}