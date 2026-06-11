namespace Convert2DTo3D
{
    public enum DisplayUnitTypeMap
    {
        // DUT_UNDEFINED,//
        DUT_CUSTOM,//

        DUT_METERS,//
        DUT_CENTIMETERS,//
        DUT_MILLIMETERS,//

        // DUT_DECIMAL_FEET,//
        DUT_FEET_FRACTIONAL_INCHES,//

        DUT_FRACTIONAL_INCHES,//

        // DUT_DECIMAL_INCHES,//
        DUT_ACRES,//

        DUT_HECTARES,//
        DUT_METERS_CENTIMETERS,//
        DUT_CUBIC_YARDS,//
        DUT_SQUARE_FEET,//
        DUT_SQUARE_METERS,//
        DUT_CUBIC_FEET,//
        DUT_CUBIC_METERS,//

        // DUT_DECIMAL_DEGREES,//
        DUT_DEGREES_AND_MINUTES,//

        DUT_GENERAL,//
        DUT_FIXED,//
        DUT_PERCENTAGE,//
        DUT_SQUARE_INCHES,//
        DUT_SQUARE_CENTIMETERS,//
        DUT_SQUARE_MILLIMETERS,//
        DUT_CUBIC_INCHES,//
        DUT_CUBIC_CENTIMETERS,//
        DUT_CUBIC_MILLIMETERS,//
        DUT_LITERS,//liter (L)
        DUT_GALLONS_US,//gallon (U.S.) (gal)
        DUT_KILOGRAMS_PER_CUBIC_METER,//kilogram per cubic meter (kg/mÂ³)
        DUT_POUNDS_MASS_PER_CUBIC_FOOT,//pound per cubic foot (lb/ftÂ³)
        DUT_POUNDS_MASS_PER_CUBIC_INCH,//pound per cubic inch (lb/inÂ³)
        DUT_BRITISH_THERMAL_UNITS,//British thermal unit[IT] (Btu[IT])
        DUT_CALORIES,//calorie[IT] (cal[IT])
        DUT_KILOCALORIES,//kilocalorie[IT] (kcal[IT])
        DUT_JOULES,//joule (J)
        DUT_KILOWATT_HOURS,//kilowatt hour (kW Â· h)
        DUT_THERMS,//therm (EC)
        DUT_INCHES_OF_WATER_PER_100FT,//Inches of water per 100 feet
        DUT_PASCALS_PER_METER,//pascal per meter (N/m)
        DUT_WATTS,//watt (W)
        DUT_KILOWATTS,//kilowatt (kW)
        DUT_BRITISH_THERMAL_UNITS_PER_SECOND,//British thermal unit[IT] per second (Btu[IT]/s)
        DUT_BRITISH_THERMAL_UNITS_PER_HOUR,//British thermal unit[IT] per hour (Btu[IT]/h)
        DUT_CALORIES_PER_SECOND,//calorie[IT] per second (cal[IT]/s)
        DUT_KILOCALORIES_PER_SECOND,//kilocalorie[IT] per second (kcal[IT]/s)
        DUT_WATTS_PER_SQUARE_FOOT,//watt per square foot (W/ftÂ²)
        DUT_WATTS_PER_SQUARE_METER,//watt per square meter (W/mÂ²)
        DUT_INCHES_OF_WATER,//inch of water (60.8Â°F)
        DUT_PASCALS,//pascal (Pa)
        DUT_KILOPASCALS,//kilopascal (kPa)
        DUT_MEGAPASCALS,//megapascal (MPa)
        DUT_POUNDS_FORCE_PER_SQUARE_INCH,//pound-force per square inch (psi) (lbf/in2)
        DUT_INCHES_OF_MERCURY,//inch of mercury conventional (inHg)
        DUT_MILLIMETERS_OF_MERCURY,//millimeter of mercury conventional (mmHg)
        DUT_ATMOSPHERES,//atmosphere standard (atm)
        DUT_BARS,//bar (bar)
        DUT_FAHRENHEIT,//degree Fahrenheit (Â°F)
        DUT_CELSIUS,//degree Celsius (Â°C)
        DUT_KELVIN,//kelvin (K)
        DUT_RANKINE,//degree Rankine (Â°R)
        DUT_FEET_PER_MINUTE,//foot per minute (ft/min)
        DUT_METERS_PER_SECOND,//meter per second (m/s)
        DUT_CENTIMETERS_PER_MINUTE,//centimeter per minute (cm/min)
        DUT_CUBIC_FEET_PER_MINUTE,//cubic foot per minute (ftÂ³/min)
        DUT_LITERS_PER_SECOND,//liter per second (L/s)
        DUT_CUBIC_METERS_PER_SECOND,//cubic meter per second (mÂ³/s)
        DUT_CUBIC_METERS_PER_HOUR,//cubic meters per hour (mÂ³/h)
        DUT_GALLONS_US_PER_MINUTE,//gallon (U.S.) per minute (gpm) (gal/min)
        DUT_GALLONS_US_PER_HOUR,//gallon (U.S.) per hour (gph) (gal/h)
        DUT_AMPERES,//ampere (A)
        DUT_KILOAMPERES,//kiloampere (kA)
        DUT_MILLIAMPERES,//milliampere (mA)
        DUT_VOLTS,//volt (V)
        DUT_KILOVOLTS,//kilovolt (kV)
        DUT_MILLIVOLTS,//millivolt (mV)
        DUT_HERTZ,//hertz (Hz)
        DUT_CYCLES_PER_SECOND,//
        DUT_LUX,//lux (lx)
        DUT_FOOTCANDLES,//footcandle
        DUT_FOOTLAMBERTS,//footlambert
        DUT_CANDELAS_PER_SQUARE_METER,//candela per square meter (cd/mÂ²)
        DUT_CANDELAS,//candela (cd)

        // DUT_CANDLEPOWER,//obsolete
        DUT_LUMENS,//lumen (lm)

        DUT_VOLT_AMPERES,//
        DUT_KILOVOLT_AMPERES,//
        DUT_HORSEPOWER,//horsepower (550 ft Â· lbf/s)
        DUT_NEWTONS,//
        DUT_DECANEWTONS,//
        DUT_KILONEWTONS,//
        DUT_MEGANEWTONS,//
        DUT_KIPS,//
        DUT_KILOGRAMS_FORCE,//
        DUT_TONNES_FORCE,//
        DUT_POUNDS_FORCE,//
        DUT_NEWTONS_PER_METER,//
        DUT_DECANEWTONS_PER_METER,//
        DUT_KILONEWTONS_PER_METER,//
        DUT_MEGANEWTONS_PER_METER,//
        DUT_KIPS_PER_FOOT,//
        DUT_KILOGRAMS_FORCE_PER_METER,//
        DUT_TONNES_FORCE_PER_METER,//
        DUT_POUNDS_FORCE_PER_FOOT,//
        DUT_NEWTONS_PER_SQUARE_METER,//
        DUT_DECANEWTONS_PER_SQUARE_METER,//
        DUT_KILONEWTONS_PER_SQUARE_METER,//
        DUT_MEGANEWTONS_PER_SQUARE_METER,//
        DUT_KIPS_PER_SQUARE_FOOT,//
        DUT_KILOGRAMS_FORCE_PER_SQUARE_METER,//
        DUT_TONNES_FORCE_PER_SQUARE_METER,//
        DUT_POUNDS_FORCE_PER_SQUARE_FOOT,//
        DUT_NEWTON_METERS,//
        DUT_DECANEWTON_METERS,//
        DUT_KILONEWTON_METERS,//
        DUT_MEGANEWTON_METERS,//
        DUT_KIP_FEET,//
        DUT_KILOGRAM_FORCE_METERS,//
        DUT_TONNE_FORCE_METERS,//
        DUT_POUND_FORCE_FEET,//
        DUT_METERS_PER_KILONEWTON,//
        DUT_FEET_PER_KIP,//
        DUT_SQUARE_METERS_PER_KILONEWTON,//
        DUT_SQUARE_FEET_PER_KIP,//
        DUT_CUBIC_METERS_PER_KILONEWTON,//
        DUT_CUBIC_FEET_PER_KIP,//
        DUT_INV_KILONEWTONS,//
        DUT_INV_KIPS,//
        DUT_FEET_OF_WATER_PER_100FT,//foot of water conventional (ftH2O) per 100 ft
        DUT_FEET_OF_WATER,//foot of water conventional (ftH2O)
        DUT_PASCAL_SECONDS,//pascal second (Pa Â· s)
        DUT_POUNDS_MASS_PER_FOOT_SECOND,//pound per foot second (lb/(ft Â· s))
        DUT_CENTIPOISES,//centipoise (cP)
        DUT_FEET_PER_SECOND,//foot per second (ft/s)
        DUT_KIPS_PER_SQUARE_INCH,//
        DUT_KILONEWTONS_PER_CUBIC_METER,//kilnewtons per cubic meter (kN/mÂ³)
        DUT_POUNDS_FORCE_PER_CUBIC_FOOT,//pound per cubic foot (kip/ftÂ³)
        DUT_KIPS_PER_CUBIC_INCH,//pound per cubic foot (kip/inÂ³)
        DUT_INV_FAHRENHEIT,//
        DUT_INV_CELSIUS,//
        DUT_NEWTON_METERS_PER_METER,//
        DUT_DECANEWTON_METERS_PER_METER,//
        DUT_KILONEWTON_METERS_PER_METER,//
        DUT_MEGANEWTON_METERS_PER_METER,//
        DUT_KIP_FEET_PER_FOOT,//
        DUT_KILOGRAM_FORCE_METERS_PER_METER,//
        DUT_TONNE_FORCE_METERS_PER_METER,//
        DUT_POUND_FORCE_FEET_PER_FOOT,//
        DUT_POUNDS_MASS_PER_FOOT_HOUR,//pound per foot hour (lb/(ft Â· h))
        DUT_KIPS_PER_INCH,//
        DUT_KIPS_PER_CUBIC_FOOT,//pound per cubic foot (kip/ftÂ³)
        DUT_KIP_FEET_PER_DEGREE,//
        DUT_KILONEWTON_METERS_PER_DEGREE,//
        DUT_KIP_FEET_PER_DEGREE_PER_FOOT,//
        DUT_KILONEWTON_METERS_PER_DEGREE_PER_METER,//
        DUT_WATTS_PER_SQUARE_METER_KELVIN,//watt per square meter kelvin (W/(mÂ² Â· K))
        DUT_BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT_FAHRENHEIT,//British thermal unit[IT] per hour square foot degree Fahrenheit (Btu[IT]/(h Â· ftÂ² Â· Â°F)
        DUT_CUBIC_FEET_PER_MINUTE_SQUARE_FOOT,//cubic foot per minute square foot
        DUT_LITERS_PER_SECOND_SQUARE_METER,//liter per second square meter
        DUT_RATIO_10,//
        DUT_RATIO_12,//
        DUT_SLOPE_DEGREES,//
        DUT_RISE_OVER_INCHES,//
        DUT_RISE_OVER_FOOT,//
        DUT_RISE_OVER_MMS,//
        DUT_WATTS_PER_CUBIC_FOOT,//watt per cubic foot (W/mÂ³)
        DUT_WATTS_PER_CUBIC_METER,//watt per cubic meter (W/mÂ³)
        DUT_BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT,//British thermal unit[IT] per hour square foot (Btu[IT]/(h Â· ftÂ²)
        DUT_BRITISH_THERMAL_UNITS_PER_HOUR_CUBIC_FOOT,//British thermal unit[IT] per hour cubic foot (Btu[IT]/(h Â· ftÂ³)
        DUT_TON_OF_REFRIGERATION,//Ton of refrigeration (12 000 Btu[IT]/h)
        DUT_CUBIC_FEET_PER_MINUTE_CUBIC_FOOT,//cubic foot per minute cubic foot
        DUT_LITERS_PER_SECOND_CUBIC_METER,//liter per second cubic meter
        DUT_CUBIC_FEET_PER_MINUTE_TON_OF_REFRIGERATION,//cubic foot per minute ton of refrigeration
        DUT_LITERS_PER_SECOND_KILOWATTS,//liter per second kilowatt
        DUT_SQUARE_FEET_PER_TON_OF_REFRIGERATION,//square foot per ton of refrigeration
        DUT_SQUARE_METERS_PER_KILOWATTS,//square meter per kilowatt
        DUT_CURRENCY,//
        DUT_LUMENS_PER_WATT,//
        DUT_SQUARE_FEET_PER_THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR,//square foot per thousand British thermal unit[IT] per hour
        DUT_KILONEWTONS_PER_SQUARE_CENTIMETER,//
        DUT_NEWTONS_PER_SQUARE_MILLIMETER,//
        DUT_KILONEWTONS_PER_SQUARE_MILLIMETER,//
        DUT_RISE_OVER_120_INCHES,//
        DUT_1_RATIO,//
        DUT_RISE_OVER_10_FEET,//
        DUT_HOUR_SQUARE_FOOT_FAHRENHEIT_PER_BRITISH_THERMAL_UNIT,//
        DUT_SQUARE_METER_KELVIN_PER_WATT,//
        DUT_BRITISH_THERMAL_UNIT_PER_FAHRENHEIT,//
        DUT_JOULES_PER_KELVIN,//
        DUT_KILOJOULES_PER_KELVIN,//
        DUT_KILOGRAMS_MASS,//kilograms (kg)
        DUT_TONNES_MASS,//tonnes (t)
        DUT_POUNDS_MASS,//pounds (lb)
        DUT_METERS_PER_SECOND_SQUARED,//meters per second squared (m/sÂ²)
        DUT_KILOMETERS_PER_SECOND_SQUARED,//kilometers per second squared (km/sÂ²)
        DUT_INCHES_PER_SECOND_SQUARED,//inches per second squared (in/sÂ²)
        DUT_FEET_PER_SECOND_SQUARED,//feet per second squared (ft/sÂ²)
        DUT_MILES_PER_SECOND_SQUARED,//miles per second squared (mi/sÂ²)
        DUT_FEET_TO_THE_FOURTH_POWER,//feet to the fourth power (ft^4)
        DUT_INCHES_TO_THE_FOURTH_POWER,//inches to the fourth power (in^4)
        DUT_MILLIMETERS_TO_THE_FOURTH_POWER,//millimeters to the fourth power (mm^4)
        DUT_CENTIMETERS_TO_THE_FOURTH_POWER,//centimeters to the fourth power (cm^4)
        DUT_METERS_TO_THE_FOURTH_POWER,//Meters to the fourth power (m^4)
        DUT_FEET_TO_THE_SIXTH_POWER,//feet to the sixth power (ft^6)
        DUT_INCHES_TO_THE_SIXTH_POWER,//inches to the sixth power (in^6)
        DUT_MILLIMETERS_TO_THE_SIXTH_POWER,//millimeters to the sixth power (mm^6)
        DUT_CENTIMETERS_TO_THE_SIXTH_POWER,//centimeters to the sixth power (cm^6)
        DUT_METERS_TO_THE_SIXTH_POWER,//meters to the sixth power (m^6)
        DUT_SQUARE_FEET_PER_FOOT,//square feet per foot (ftÂ²/ft)
        DUT_SQUARE_INCHES_PER_FOOT,//square inches per foot (inÂ²/ft)
        DUT_SQUARE_MILLIMETERS_PER_METER,//square millimeters per meter (mmÂ²/m)
        DUT_SQUARE_CENTIMETERS_PER_METER,//square centimeters per meter (cmÂ²/m)
        DUT_SQUARE_METERS_PER_METER,//square meters per meter (mÂ²/m)
        DUT_KILOGRAMS_MASS_PER_METER,//kilograms per meter (kg/m)
        DUT_POUNDS_MASS_PER_FOOT,//pounds per foot (lb/ft)
        DUT_RADIANS,//radians
        DUT_GRADS,//grads
        DUT_RADIANS_PER_SECOND,//radians per second
        DUT_MILISECONDS,//millisecond
        DUT_SECONDS,//second
        DUT_MINUTES,//minutes
        DUT_HOURS,//hours
        DUT_KILOMETERS_PER_HOUR,//kilometers per hour
        DUT_MILES_PER_HOUR,//miles per hour
        DUT_KILOJOULES,//kilojoules
        DUT_KILOGRAMS_MASS_PER_SQUARE_METER,//kilograms per square meter (kg/mÂ²)
        DUT_POUNDS_MASS_PER_SQUARE_FOOT,//pounds per square foot (lb/ftÂ²)
        DUT_WATTS_PER_METER_KELVIN,//Watts per meter kelvin (W/(mÂ·K))
        DUT_JOULES_PER_GRAM_CELSIUS,//Joules per gram Celsius (J/(gÂ·Â°C))
        DUT_JOULES_PER_GRAM,//Joules per gram (J/g)
        DUT_NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER,//Nanograms per pascal second square meter (ng/(PaÂ·sÂ·mÂ²))
        DUT_OHM_METERS,//Ohm meters
        DUT_BRITISH_THERMAL_UNITS_PER_HOUR_FOOT_FAHRENHEIT,//BTU per hour foot Fahrenheit (BTU/(hÂ·ftÂ·Â°F))
        DUT_BRITISH_THERMAL_UNITS_PER_POUND_FAHRENHEIT,//BTU per pound Fahrenheit (BTU/(lbÂ·Â°F))
        DUT_BRITISH_THERMAL_UNITS_PER_POUND,//BTU per pound (BTU/lb)
        DUT_GRAINS_PER_HOUR_SQUARE_FOOT_INCH_MERCURY,//Grains per hour square foot inch mercury (gr/(hÂ·ftÂ²Â·inHg))
        DUT_PER_MILLE,//Per mille or per thousand(â€°)
        DUT_DECIMETERS,//Decimeters
        DUT_JOULES_PER_KILOGRAM_CELSIUS,//Joules per kilogram Celsius (J/(kgÂ·Â°C))
        DUT_MICROMETERS_PER_METER_CELSIUS,//Micrometers per meter Celsius (um/(mÂ·Â°C))
        DUT_MICROINCHES_PER_INCH_FAHRENHEIT,//Microinches per inch Fahrenheit (uin/(inÂ·Â°F))
        DUT_USTONNES_MASS,//US tonnes (T, Tons, ST)
        DUT_USTONNES_FORCE,//US tonnes (Tonsf, STf)
        DUT_LITERS_PER_MINUTE,//liters per minute (L/min)
        DUT_FAHRENHEIT_DIFFERENCE,//degree Fahrenheit difference (delta Â°F)
        DUT_CELSIUS_DIFFERENCE,//degree Celsius difference (delta Â°C)
        DUT_KELVIN_DIFFERENCE,//kelvin difference (delta K)
        DUT_RANKINE_DIFFERENCE,//degree Rankine difference (delta Â°R)
    }
}