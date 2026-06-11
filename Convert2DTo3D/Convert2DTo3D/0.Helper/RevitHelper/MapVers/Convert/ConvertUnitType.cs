using Autodesk.Revit.DB;
using System;

namespace Convert2DTo3D
{
    /*
		List<string> forceTypeNames = new List<string>();
		Type unitTypeIdType = typeof(SpecTypeId);
		PropertyInfo[] properties = unitTypeIdType.GetProperties(BindingFlags.Static | BindingFlags.Public);
		foreach (PropertyInfo property in properties) forceTypeNames.Add(property.Name);
		foreach (int i in Enum.GetValues(typeof(UnitTypeMap)))
		{
			var enumValuen = Enum.GetName(typeof(UnitTypeMap), i);
			enumValuen = enumValuen.Replace("UT", "").ToLower();
			var enumValue= enumValuen.Split('_');

			var name = string.Join("", enumValue) ;
			var chek = forceTypeNames.Where(k => k.ToLower().Contains(name) || k.ToLower() == name).ToList();

			if (chek.Count() == 0)
			{
				Debug.WriteLine($"case UnitTypeMap.{Enum.GetName(typeof(UnitTypeMap), i)}: \n\treturn SpecTypeId.Valid;");
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
				Debug.WriteLine($"case UnitTypeMap.{Enum.GetName(typeof(UnitTypeMap), i)}: \n\t  return SpecTypeId.{ne};");
			}
		}
  */

    public class ConvertUnitType
    {
#if REVIT2019 || REVIT2020

        public static UnitType Convert(UnitTypeMap unitTypeMap)
        {
            var stringEnum = unitTypeMap.ToString();
            var checkExists = Enum.IsDefined(typeof(UnitType), stringEnum);
            if (checkExists == true)
            {
                UnitType output = UnitType.UT_Length;
                Enum.TryParse<UnitType>(stringEnum, out output);
                return output;
            }
            return UnitType.UT_Length;
        }

#else

        public static ForgeTypeId Convert(UnitTypeMap unitTypeMap)
        {
            switch (unitTypeMap)
            {
                case UnitTypeMap.UT_Custom:
                    return SpecTypeId.Custom;

                case UnitTypeMap.UT_Length:
                    return SpecTypeId.Length;

                case UnitTypeMap.UT_Area:
                    return SpecTypeId.Area;

                case UnitTypeMap.UT_Volume:
                    return SpecTypeId.Volume;

                case UnitTypeMap.UT_Angle:
                    return SpecTypeId.Angle;

                case UnitTypeMap.UT_Number:
                    return SpecTypeId.Number;

                case UnitTypeMap.UT_SheetLength:
                    return SpecTypeId.SheetLength;

                case UnitTypeMap.UT_SiteAngle:
                    return SpecTypeId.SiteAngle;

                case UnitTypeMap.UT_HVAC_Density:
                    return SpecTypeId.HvacDensity;

                case UnitTypeMap.UT_HVAC_Energy:
                    return SpecTypeId.HvacEnergy;

                case UnitTypeMap.UT_HVAC_Friction:
                    return SpecTypeId.HvacFriction;

                case UnitTypeMap.UT_HVAC_Power:
                    return SpecTypeId.HvacPower;

                case UnitTypeMap.UT_HVAC_Power_Density:
                    return SpecTypeId.HvacPowerDensity;

                case UnitTypeMap.UT_HVAC_Pressure:
                    return SpecTypeId.HvacPressure;

                case UnitTypeMap.UT_HVAC_Temperature:
                    return SpecTypeId.HvacTemperature;

                case UnitTypeMap.UT_HVAC_Velocity:
                    return SpecTypeId.HvacVelocity;

                case UnitTypeMap.UT_HVAC_Airflow:
                    return SpecTypeId.AirFlow;

                case UnitTypeMap.UT_HVAC_DuctSize:
                    return SpecTypeId.DuctSize;

                case UnitTypeMap.UT_HVAC_CrossSection:
                    return SpecTypeId.CrossSection;

                case UnitTypeMap.UT_HVAC_HeatGain:
                    return SpecTypeId.HeatGain;

                case UnitTypeMap.UT_Electrical_Current:
                    return SpecTypeId.Current; //
                case UnitTypeMap.UT_Electrical_Potential:
                    return SpecTypeId.ElectricalPotential;

                case UnitTypeMap.UT_Electrical_Frequency:
                    return SpecTypeId.ElectricalFrequency;

                case UnitTypeMap.UT_Electrical_Illuminance:
                    return SpecTypeId.Illuminance; //
                case UnitTypeMap.UT_Electrical_Luminous_Flux:
                    return SpecTypeId.LuminousFlux;  //
                case UnitTypeMap.UT_Electrical_Power:
                    return SpecTypeId.ElectricalPower;

                case UnitTypeMap.UT_HVAC_Roughness:
                    return SpecTypeId.HvacRoughness;

                case UnitTypeMap.UT_Force:
                    return SpecTypeId.Force;

                case UnitTypeMap.UT_LinearForce:
                    return SpecTypeId.LinearForce;

                case UnitTypeMap.UT_AreaForce:
                    return SpecTypeId.AreaForce;

                case UnitTypeMap.UT_Moment:
                    return SpecTypeId.Moment;

                case UnitTypeMap.UT_ForceScale:
                    return SpecTypeId.ForceScale;

                case UnitTypeMap.UT_LinearForceScale:
                    return SpecTypeId.LinearForceScale;

                case UnitTypeMap.UT_AreaForceScale:
                    return SpecTypeId.AreaForceScale;

                case UnitTypeMap.UT_MomentScale:
                    return SpecTypeId.MomentScale;

                case UnitTypeMap.UT_Electrical_Apparent_Power:
                    return SpecTypeId.ApparentPower;

                case UnitTypeMap.UT_Electrical_Power_Density:
                    return SpecTypeId.ElectricalPowerDensity;

                case UnitTypeMap.UT_Piping_Density:
                    return SpecTypeId.PipingDensity;

                case UnitTypeMap.UT_Piping_Flow:
                    return SpecTypeId.Flow;

                case UnitTypeMap.UT_Piping_Friction:
                    return SpecTypeId.PipingFriction;

                case UnitTypeMap.UT_Piping_Pressure:
                    return SpecTypeId.PipingPressure;

                case UnitTypeMap.UT_Piping_Temperature:
                    return SpecTypeId.PipingTemperature;

                case UnitTypeMap.UT_Piping_Velocity:
                    return SpecTypeId.PipingVelocity;

                case UnitTypeMap.UT_Piping_Viscosity:
                    return SpecTypeId.PipingViscosity;

                case UnitTypeMap.UT_PipeSize:
                    return SpecTypeId.PipeSize;

                case UnitTypeMap.UT_Piping_Roughness:
                    return SpecTypeId.PipingRoughness;

                case UnitTypeMap.UT_Stress:
                    return SpecTypeId.Stress;

                case UnitTypeMap.UT_UnitWeight:
                    return SpecTypeId.UnitWeight;

                case UnitTypeMap.UT_ThermalExpansion:
                    return SpecTypeId.ThermalExpansionCoefficient;

                case UnitTypeMap.UT_LinearMoment:
                    return SpecTypeId.LinearMoment;

                case UnitTypeMap.UT_LinearMomentScale:
                    return SpecTypeId.LinearMomentScale;

                case UnitTypeMap.UT_Piping_Volume:
                    return SpecTypeId.PipingVolume;

                case UnitTypeMap.UT_HVAC_Viscosity:
                    return SpecTypeId.HvacViscosity;

                case UnitTypeMap.UT_HVAC_CoefficientOfHeatTransfer:
                    return SpecTypeId.HeatTransferCoefficient;

                case UnitTypeMap.UT_HVAC_Airflow_Density:
                    return SpecTypeId.AirFlowDensity;

                case UnitTypeMap.UT_Slope:
                    return SpecTypeId.Slope;

                case UnitTypeMap.UT_HVAC_Cooling_Load:
                    return SpecTypeId.CoolingLoad;

                case UnitTypeMap.UT_HVAC_Cooling_Load_Divided_By_Area:
                    return SpecTypeId.CoolingLoadDividedByArea;

                case UnitTypeMap.UT_HVAC_Cooling_Load_Divided_By_Volume:
                    return SpecTypeId.CoolingLoadDividedByVolume;

                case UnitTypeMap.UT_HVAC_Heating_Load:
                    return SpecTypeId.HeatingLoad;

                case UnitTypeMap.UT_HVAC_Heating_Load_Divided_By_Area:
                    return SpecTypeId.HeatingLoadDividedByArea;

                case UnitTypeMap.UT_HVAC_Heating_Load_Divided_By_Volume:
                    return SpecTypeId.HeatingLoadDividedByVolume;

                case UnitTypeMap.UT_HVAC_Airflow_Divided_By_Volume:
                    return SpecTypeId.AirFlowDividedByVolume;

                case UnitTypeMap.UT_HVAC_Airflow_Divided_By_Cooling_Load:
                    return SpecTypeId.AirFlowDividedByCoolingLoad;

                case UnitTypeMap.UT_HVAC_Area_Divided_By_Cooling_Load:
                    return SpecTypeId.AreaDividedByCoolingLoad;

                case UnitTypeMap.UT_HVAC_Slope:
                    return SpecTypeId.HvacSlope;

                case UnitTypeMap.UT_Piping_Slope:
                    return SpecTypeId.PipingSlope;

                case UnitTypeMap.UT_Currency:
                    return SpecTypeId.Currency;

                case UnitTypeMap.UT_Electrical_Efficacy:
                    return SpecTypeId.Efficacy;

                case UnitTypeMap.UT_Electrical_Wattage:
                    return SpecTypeId.Wattage;

                case UnitTypeMap.UT_Color_Temperature:
                    return SpecTypeId.ColorTemperature;

                case UnitTypeMap.UT_DecSheetLength:
                    return SpecTypeId.DecimalSheetLength;

                case UnitTypeMap.UT_Electrical_Luminous_Intensity:
                    return SpecTypeId.LuminousIntensity;

                case UnitTypeMap.UT_Electrical_Luminance:
                    return SpecTypeId.Luminance;

                case UnitTypeMap.UT_HVAC_Area_Divided_By_Heating_Load:
                    return SpecTypeId.AreaDividedByHeatingLoad;

                case UnitTypeMap.UT_HVAC_Factor:
                    return SpecTypeId.Factor;

                case UnitTypeMap.UT_Electrical_Temperature:
                    return SpecTypeId.ElectricalTemperature;

                case UnitTypeMap.UT_Electrical_CableTraySize:
                    return SpecTypeId.CableTraySize;

                case UnitTypeMap.UT_Electrical_ConduitSize:
                    return SpecTypeId.ConduitSize;

                case UnitTypeMap.UT_Reinforcement_Volume:
                    return SpecTypeId.ReinforcementVolume;

                case UnitTypeMap.UT_Reinforcement_Length:
                    return SpecTypeId.ReinforcementLength;

                case UnitTypeMap.UT_Electrical_Demand_Factor:
                    return SpecTypeId.DemandFactor;

                case UnitTypeMap.UT_HVAC_DuctInsulationThickness:
                    return SpecTypeId.DuctInsulationThickness;

                case UnitTypeMap.UT_HVAC_DuctLiningThickness:
                    return SpecTypeId.DuctLiningThickness;

                case UnitTypeMap.UT_PipeInsulationThickness:
                    return SpecTypeId.PipeInsulationThickness;

                case UnitTypeMap.UT_HVAC_ThermalResistance:
                    return SpecTypeId.ThermalResistance;

                case UnitTypeMap.UT_HVAC_ThermalMass:
                    return SpecTypeId.ThermalMass;

                case UnitTypeMap.UT_Acceleration:
                    return SpecTypeId.Acceleration;

                case UnitTypeMap.UT_Bar_Diameter:
                    return SpecTypeId.BarDiameter;

                case UnitTypeMap.UT_Crack_Width:
                    return SpecTypeId.CrackWidth;

                case UnitTypeMap.UT_Displacement_Deflection:
                    return SpecTypeId.Displacement;

                case UnitTypeMap.UT_Energy:
                    return SpecTypeId.Energy;

                case UnitTypeMap.UT_Structural_Frequency:
                    return SpecTypeId.StructuralFrequency;

                case UnitTypeMap.UT_Mass:
                    return SpecTypeId.Mass;

                case UnitTypeMap.UT_Mass_per_Unit_Length:
                    return SpecTypeId.MassPerUnitLength;

                case UnitTypeMap.UT_Moment_of_Inertia:
                    return SpecTypeId.MomentOfInertia;

                case UnitTypeMap.UT_Surface_Area:
                    return SpecTypeId.SurfaceAreaPerUnitLength;

                case UnitTypeMap.UT_Period:
                    return SpecTypeId.Period;

                case UnitTypeMap.UT_Pulsation:
                    return SpecTypeId.Pulsation;

                case UnitTypeMap.UT_Reinforcement_Area:
                    return SpecTypeId.ReinforcementArea;

                case UnitTypeMap.UT_Reinforcement_Area_per_Unit_Length:
                    return SpecTypeId.ReinforcementAreaPerUnitLength;

                case UnitTypeMap.UT_Reinforcement_Cover:
                    return SpecTypeId.ReinforcementCover;

                case UnitTypeMap.UT_Reinforcement_Spacing:
                    return SpecTypeId.ReinforcementSpacing;

                case UnitTypeMap.UT_Rotation:
                    return SpecTypeId.Rotation;

                case UnitTypeMap.UT_Section_Area:
                    return SpecTypeId.SectionArea;

                case UnitTypeMap.UT_Section_Dimension:
                    return SpecTypeId.SectionDimension;

                case UnitTypeMap.UT_Section_Modulus:
                    return SpecTypeId.SectionModulus;

                case UnitTypeMap.UT_Section_Property:
                    return SpecTypeId.SectionProperty;

                case UnitTypeMap.UT_Structural_Velocity:
                    return SpecTypeId.StructuralVelocity;

                case UnitTypeMap.UT_Warping_Constant:
                    return SpecTypeId.WarpingConstant;

                case UnitTypeMap.UT_Weight:
                    return SpecTypeId.Weight;

                case UnitTypeMap.UT_Weight_per_Unit_Length:
                    return SpecTypeId.WeightPerUnitLength;

                case UnitTypeMap.UT_HVAC_ThermalConductivity:
                    return SpecTypeId.ThermalConductivity;

                case UnitTypeMap.UT_HVAC_SpecificHeat:
                    return SpecTypeId.SpecificHeat;

                case UnitTypeMap.UT_HVAC_SpecificHeatOfVaporization:
                    return SpecTypeId.SpecificHeatOfVaporization;

                case UnitTypeMap.UT_HVAC_Permeability:
                    return SpecTypeId.Permeability;

                case UnitTypeMap.UT_Electrical_Resistivity:
                    return SpecTypeId.ElectricalResistivity;

                case UnitTypeMap.UT_MassDensity:
                    return SpecTypeId.MassDensity;

                case UnitTypeMap.UT_MassPerUnitArea:
                    return SpecTypeId.MassPerUnitArea;

                case UnitTypeMap.UT_Pipe_Dimension:
                    return SpecTypeId.PipeDimension;

                case UnitTypeMap.UT_PipeMass:
                    return SpecTypeId.PipeMassPerUnitLength;

                case UnitTypeMap.UT_PipeMassPerUnitLength:
                    return SpecTypeId.PipeMassPerUnitLength;

                case UnitTypeMap.UT_HVAC_TemperatureDifference:
                    return SpecTypeId.HvacTemperatureDifference;

                case UnitTypeMap.UT_Piping_TemperatureDifference:
                    return SpecTypeId.PipingTemperatureDifference;

                case UnitTypeMap.UT_Electrical_TemperatureDifference:
                    return SpecTypeId.ElectricalTemperatureDifference;

                case UnitTypeMap.UT_TimeInterval:
                    return SpecTypeId.Time;

                case UnitTypeMap.UT_Speed:
                    return SpecTypeId.Speed;

                default:
                    return SpecTypeId.Length;
            }
        }

#endif
    }
}