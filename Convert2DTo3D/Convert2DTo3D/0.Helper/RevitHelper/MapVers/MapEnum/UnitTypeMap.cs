namespace Convert2DTo3D
{
    public enum UnitTypeMap
    {
        // UT_Undefined,// Undefined unit value
        UT_Custom,// A custom unit value

        UT_Length,// Length, e.g. ft, in, m, mm
        UT_Area,// Area, e.g. ftÂ², inÂ², mÂ², mmÂ²
        UT_Volume,// Volume, e.g. ftÂ³, inÂ³, mÂ³, mmÂ³
        UT_Angle,// Angular measurement, e.g. radians, degrees
        UT_Number,// General format unit, appropriate for general counts or percentages
        UT_SheetLength,// Sheet length
        UT_SiteAngle,// Site angle
        UT_HVAC_Density,// Density (HVAC) e.g. kg/mÂ³
        UT_HVAC_Energy,// Energy (HVAC) e.g. (mÂ² Â· kg)/sÂ², J
        UT_HVAC_Friction,// Friction (HVAC) e.g. kg/(mÂ² Â· sÂ²), Pa/m
        UT_HVAC_Power,// Power (HVAC) e.g. (mÂ² Â· kg)/sÂ³, W
        UT_HVAC_Power_Density,// Power Density (HVAC), e.g. kg/sÂ³, W/mÂ²
        UT_HVAC_Pressure,// Pressure (HVAC) e.g. kg/(m Â· sÂ²), Pa
        UT_HVAC_Temperature,// Temperature (HVAC) e.g. K, C, F
        UT_HVAC_Velocity,// Velocity (HVAC) e.g. m/s
        UT_HVAC_Airflow,// Air Flow (HVAC) e.g. mÂ³/s
        UT_HVAC_DuctSize,// Duct Size (HVAC) e.g. mm, in
        UT_HVAC_CrossSection,// Cross Section (HVAC) e.g. mmÂ², inÂ²
        UT_HVAC_HeatGain,// Heat Gain (HVAC) e.g. (mÂ² Â· kg)/sÂ³, W
        UT_Electrical_Current,// Current (Electrical) e.g. A
        UT_Electrical_Potential,// Electrical Potential e.g. (mÂ² Â· kg) / (sÂ³Â· A), V
        UT_Electrical_Frequency,// Frequency (Electrical) e.g. 1/s, Hz
        UT_Electrical_Illuminance,// Illuminance (Electrical) e.g. (cd Â· sr)/mÂ², lm/mÂ²
        UT_Electrical_Luminous_Flux,// Luminous Flux (Electrical) e.g. cd Â· sr, lm
        UT_Electrical_Power,// Power (Electrical) e.g. (mÂ² Â· kg)/sÂ³, W
        UT_HVAC_Roughness,// Roughness factor (HVAC) e,g. ft, in, mm
        UT_Force,// Force, e.g. (kg Â· m)/sÂ², N
        UT_LinearForce,// Force per unit length, e.g. kg/sÂ², N/m
        UT_AreaForce,// Force per unit area, e.g. kg/(m Â· sÂ²), N/mÂ²
        UT_Moment,// Moment, e.g. (kg Â· mÂ²)/sÂ², N Â· m
        UT_ForceScale,// Force scale, e.g. m / N
        UT_LinearForceScale,// Linear force scale, e.g. mÂ² / N
        UT_AreaForceScale,// Area force scale, e.g. mÂ³ / N
        UT_MomentScale,// Moment scale, e.g. 1 / N
        UT_Electrical_Apparent_Power,// Apparent Power (Electrical), e.g. (mÂ² Â· kg)/sÂ³, W
        UT_Electrical_Power_Density,// Power Density (Electrical), e.g. kg/sÂ³, W/mÂ²
        UT_Piping_Density,// Density (Piping) e.g. kg/mÂ³
        UT_Piping_Flow,// Flow (Piping), e.g. mÂ³/s
        UT_Piping_Friction,// Friction (Piping), e.g. kg/(mÂ² Â· sÂ²), Pa/m
        UT_Piping_Pressure,// Pressure (Piping), e.g. kg/(m Â· sÂ²), Pa
        UT_Piping_Temperature,// Temperature (Piping), e.g. K
        UT_Piping_Velocity,// Velocity (Piping), e.g. m/s
        UT_Piping_Viscosity,// Dynamic Viscosity (Piping), e.g. kg/(m Â· s), Pa Â· s
        UT_PipeSize,// Pipe Size (Piping), e.g. m
        UT_Piping_Roughness,// Roughness factor (Piping), e.g. ft, in, mm
        UT_Stress,// Stress, e.g. kg/(m Â· sÂ²), ksi, MPa
        UT_UnitWeight,// Unit weight, e.g. N/mÂ³
        UT_ThermalExpansion,// Thermal expansion, e.g. 1/K
        UT_LinearMoment,// Linear moment, e,g. (N Â· m)/m, lbf / ft
        UT_LinearMomentScale,// Linear moment scale, e.g. ft/kip, m/kN

        // UT_ForcePerLength,// Point Spring Coefficient, e.g. kg/sÂ², N/m
        // UT_ForceLengthPerAngle,// Rotational Point Spring Coefficient, e.g. (kg Â· mÂ²)/(sÂ² Â· rad), (N Â· m)/rad
        //UT_LinearForcePerLength,// Line Spring Coefficient, e.g. kg/(m Â· sÂ²), (N Â· m)/mÂ²
        //UT_LinearForceLengthPerAngle,// Rotational Line Spring Coefficient, e.g. (kg Â· m)/(sÂ² Â· rad), N/rad
        //UT_AreaForcePerLength,// Area Spring Coefficient, e.g. kg/(mÂ² Â· sÂ²), N/mÂ³
        UT_Piping_Volume,// Pipe Volume, e.g. gallons, liters

        UT_HVAC_Viscosity,// Dynamic Viscosity (HVAC), e.g. kg/(m Â· s), Pa Â· s
        UT_HVAC_CoefficientOfHeatTransfer,// Coefficient of Heat Transfer (U-value) (HVAC), e.g. kg/(sÂ³ Â· K), W/(mÂ² Â· K)
        UT_HVAC_Airflow_Density,// Air Flow Density (HVAC), mÂ³/(s Â· mÂ²)
        UT_Slope,// Slope, rise/run
        UT_HVAC_Cooling_Load,// Cooling load (HVAC), e.g. (mÂ² Â· kg)/sÂ³, W, kW, Btu/s, Btu/h
        UT_HVAC_Cooling_Load_Divided_By_Area,// Cooling load per unit area (HVAC), e.g. kg/sÂ³, W/mÂ², W/ftÂ², Btu/(hÂ·ftÂ²)
        UT_HVAC_Cooling_Load_Divided_By_Volume,// Cooling load per unit volume (HVAC), e.g. kg/(sÂ³ Â· m), W/mÂ³, Btu/(hÂ·ftÂ³)
        UT_HVAC_Heating_Load,// Heating load (HVAC), e.g. (mÂ² Â· kg)/sÂ³, W, kW, Btu/s, Btu/h
        UT_HVAC_Heating_Load_Divided_By_Area,// Heating load per unit area (HVAC), e.g. kg/sÂ³, W/mÂ², W/ftÂ², Btu/(hÂ·ftÂ²)
        UT_HVAC_Heating_Load_Divided_By_Volume,// Heating load per unit volume (HVAC), e.g. kg/(sÂ³ Â· m), W/mÂ³, Btu/(hÂ·ftÂ³)
        UT_HVAC_Airflow_Divided_By_Volume,// Airflow per unit volume (HVAC), e.g. mÂ³/(s Â· mÂ³), CFM/ftÂ³, CFM/CF, L/(sÂ·mÂ³)
        UT_HVAC_Airflow_Divided_By_Cooling_Load,// Airflow per unit cooling load (HVAC), e.g. (m Â· sÂ²)/kg, ftÂ²/ton, SF/ton, mÂ²/kW
        UT_HVAC_Area_Divided_By_Cooling_Load,// Area per unit cooling load (HVAC), e.g. sÂ³/kg, ftÂ²/ton, mÂ²/kW

        //UT_WireSize,// Wire Size (Electrical), e.g. mm, inch
        UT_HVAC_Slope,// Slope (HVAC)

        UT_Piping_Slope,// Slope (Piping)
        UT_Currency,// Currency
        UT_Electrical_Efficacy,// Electrical efficacy (lighting), e.g. cdÂ·srÂ·sÂ³/(mÂ²Â·kg), lm/W
        UT_Electrical_Wattage,// Wattage (lighting), e.g. (mÂ² Â· kg)/sÂ³, W
        UT_Color_Temperature,// Color temperature (lighting), e.g. K
        UT_DecSheetLength,// Sheet length in decimal form, decimal inches, mm
        UT_Electrical_Luminous_Intensity,// Luminous Intensity (Lighting), e.g. cd, cd
        UT_Electrical_Luminance,// Luminance (Lighting), cd/mÂ², cd/mÂ²
        UT_HVAC_Area_Divided_By_Heating_Load,// Area per unit heating load (HVAC), e.g. sÂ³/kg, ftÂ²/ton, mÂ²/kW
        UT_HVAC_Factor,// Heating and coooling factor, percentage
        UT_Electrical_Temperature,// Temperature (electrical), e.g. F, C
        UT_Electrical_CableTraySize,// Cable tray size (electrical), e.g. in, mm
        UT_Electrical_ConduitSize,// Conduit size (electrical), e.g. in, mm
        UT_Reinforcement_Volume,// Structural reinforcement volume, e.g. inÂ³, cmÂ³
        UT_Reinforcement_Length,// Structural reinforcement length, e.g. mm, in, ft
        UT_Electrical_Demand_Factor,// Electrical demand factor, percentage
        UT_HVAC_DuctInsulationThickness,// Duct Insulation Thickness (HVAC), e.g. mm, in
        UT_HVAC_DuctLiningThickness,// Duct Lining Thickness (HVAC), e.g. mm, in
        UT_PipeInsulationThickness,// Pipe Insulation Thickness (Piping), e.g. mm, in
        UT_HVAC_ThermalResistance,// Thermal Resistance (HVAC), R Value, e.g. mÂ²Â·K/W
        UT_HVAC_ThermalMass,// Thermal Mass (HVAC), e.g. J/K, BTU/F
        UT_Acceleration,// Acceleration, e.g. m/sÂ², km/sÂ², in/sÂ², ft/sÂ², mi/sÂ²
        UT_Bar_Diameter,// Bar Diameter, e.g. ', LF, ", m, cm, mm
        UT_Crack_Width,// Crack Width, e.g. ', LF, ", m, cm, mm
        UT_Displacement_Deflection,// Displacement/Deflection, e.g. ', LF, ", m, cm, mm
        UT_Energy,// Energy, e.g. J, kJ, kgf-m, lb-ft, N-m
        UT_Structural_Frequency,// FREQUENCY, Frequency (Structural) e.g. Hz
        UT_Mass,// Mass, e.g. kg, lb, t
        UT_Mass_per_Unit_Length,// Mass per Unit Length, e.g. kg/m, lb/ft
        UT_Moment_of_Inertia,// Moment of Inertia, e.g. ft^4, in^4, mm^4, cm^4, m^4
        UT_Surface_Area,// Surface Area, e.g. ftÂ²/ft, mÂ²/m
        UT_Period,// Period, e.g. ms, s, min, h
        UT_Pulsation,// Pulsation, e.g. rad/s
        UT_Reinforcement_Area,// Reinforcement Area, e.g. SF, ftÂ², inÂ², mmÂ², cmÂ², mÂ²
        UT_Reinforcement_Area_per_Unit_Length,// Reinforcement Area per Unit Length, e.g. ftÂ²/ft, inÂ²/ft, mmÂ²/m, cmÂ²/m, mÂ²/m
        UT_Reinforcement_Cover,// Reinforcement Cover, e.g. ', LF, ", m, cm, mm
        UT_Reinforcement_Spacing,// Reinforcement Spacing, e.g. ', LF, ", m, cm, mm
        UT_Rotation,// Rotation, e.g. Â°, rad, grad
        UT_Section_Area,// Section Area, e.g. ftÂ²/ft, inÂ²/ft, mmÂ²/m, cmÂ²/m, mÂ²/m
        UT_Section_Dimension,// Section Dimension, e.g. ', LF, ", m, cm, mm
        UT_Section_Modulus,// Section Modulus, e.g. ft^3, in^3, mm^3, cm^3, m^3
        UT_Section_Property,// Section Property, e.g. ', LF, ", m, cm, mm
        UT_Structural_Velocity,// Section Property, e.g. km/h, m/s, ft/min, ft/s, mph
        UT_Warping_Constant,// Warping Constant, e.g. ft^6, in^6, mm^6, cm^6, m^6
        UT_Weight,// Weight, e.g. N, daN, kN, MN, kip, kgf, Tf, lb, lbf
        UT_Weight_per_Unit_Length,// Weight per Unit Length, e.g. N/m, daN/m, kN/m, MN/m, kip/ft, kgf/m, Tf/m, lb/ft, lbf/ft, kip/in
        UT_HVAC_ThermalConductivity,// Thermal Conductivity (HVAC), e.g. W/(mÂ·K)
        UT_HVAC_SpecificHeat,// Specific Heat (HVAC), e.g. J/(gÂ·Â°C)
        UT_HVAC_SpecificHeatOfVaporization,// Specific Heat of Vaporization, e.g. J/g
        UT_HVAC_Permeability,// Permeability, e.g. ng/(PaÂ·sÂ·mÂ²)
        UT_Electrical_Resistivity,// Electrical Resistivity, e.g.
        UT_MassDensity,// Mass Density, e.g. kg/mÂ³, lb/ftÂ³
        UT_MassPerUnitArea,// Mass Per Unit Area, e.g. kg/mÂ², lb/ftÂ²
        UT_Pipe_Dimension,// Length unit for pipe dimension, e.g. in, mm
        UT_PipeMass,// Mass, e.g. kg, lb, t
        UT_PipeMassPerUnitLength,// Mass per Unit Length, e.g. kg/m, lb/ft
        UT_HVAC_TemperatureDifference,// Temperature Difference (HVAC) e.g. C, F, K, R
        UT_Piping_TemperatureDifference,// Temperature Difference (Piping), e.g. C, F, K, R
        UT_Electrical_TemperatureDifference,// Temperature Difference (Electrical), e.g. C, F, K, R
        UT_TimeInterval,// Interval of time e.g. ms, s, min, h
        UT_Speed,// Distance interval over time e.g. m/h etc.
    }
}