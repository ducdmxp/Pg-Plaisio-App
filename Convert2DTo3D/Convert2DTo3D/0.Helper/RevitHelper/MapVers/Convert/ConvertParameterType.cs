using Autodesk.Revit.DB;
using System;

namespace Convert2DTo3D
{
    public class ConvertParameterType
    {
#if (REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021)

        public static ParameterType Convert(ParameterTypeMap parameterTypeMap)
        {
            var stringEnum = parameterTypeMap.ToString();
            var checkExists = Enum.IsDefined(typeof(ParameterType), stringEnum);
            if (checkExists == true)
            {
                ParameterType output = ParameterType.Text;
                Enum.TryParse<ParameterType>(stringEnum, out output);
                return output;
            }
            return ParameterType.Text;
        }

#else

        public static ForgeTypeId Convert(ParameterTypeMap ParameterTypeMap)
        {
            switch (ParameterTypeMap)
            {
                case ParameterTypeMap.Text:
                    return SpecTypeId.String.Text;

                case ParameterTypeMap.Integer:
                    return SpecTypeId.Int.Integer;

                case ParameterTypeMap.Number:
                    return SpecTypeId.Number;

                case ParameterTypeMap.Length:
                    return SpecTypeId.Length;

                case ParameterTypeMap.Area:
                    return SpecTypeId.Area;

                case ParameterTypeMap.Volume:
                    return SpecTypeId.Volume;

                case ParameterTypeMap.Angle:
                    return SpecTypeId.Angle;

                case ParameterTypeMap.URL:
                    return SpecTypeId.String.Url;

                case ParameterTypeMap.Material:
                    return SpecTypeId.Reference.Material;

                case ParameterTypeMap.YesNo:
                    return SpecTypeId.Boolean.YesNo;

                case ParameterTypeMap.Force:
                    return SpecTypeId.Force;

                case ParameterTypeMap.LinearForce:
                    return SpecTypeId.LinearForce;

                case ParameterTypeMap.AreaForce:
                    return SpecTypeId.AreaForce;

                case ParameterTypeMap.Moment:
                    return SpecTypeId.Moment;

                case ParameterTypeMap.NumberOfPoles:
                    return SpecTypeId.Int.NumberOfPoles;
                //case ParameterTypeMap.FamilyType:
                //    return SpecTypeId.FamilyType;
                case ParameterTypeMap.LoadClassification:
                    return SpecTypeId.Reference.LoadClassification;

                case ParameterTypeMap.Image:
                    return SpecTypeId.Reference.Image;

                case ParameterTypeMap.MultilineText:
                    return SpecTypeId.String.MultilineText;

                case ParameterTypeMap.HVACDensity:
                    return SpecTypeId.HvacDensity;

                case ParameterTypeMap.HVACEnergy:
                    return SpecTypeId.HvacEnergy;

                case ParameterTypeMap.HVACFriction:
                    return SpecTypeId.HvacFriction;

                case ParameterTypeMap.HVACPower:
                    return SpecTypeId.HvacPower;

                case ParameterTypeMap.HVACPowerDensity:
                    return SpecTypeId.HvacPowerDensity;

                case ParameterTypeMap.HVACPressure:
                    return SpecTypeId.HvacPressure;

                case ParameterTypeMap.HVACTemperature:
                    return SpecTypeId.HvacTemperature;

                case ParameterTypeMap.HVACVelocity:
                    return SpecTypeId.HvacVelocity;
                //case ParameterTypeMap.HVACAirflow:
                //    return SpecTypeId.FixtureUnit;
                //case ParameterTypeMap.HVACDuctSize:
                //    return SpecTypeId.FixtureUnit;
                //case ParameterTypeMap.HVACCrossSection:
                //    return SpecTypeId.FixtureUnit;
                //case ParameterTypeMap.HVACHeatGain:
                //    return SpecTypeId.FixtureUnit;
                //case ParameterTypeMap.ElectricalCurrent:
                //    return SpecTypeId.FixtureUnit;
                case ParameterTypeMap.ElectricalPotential:
                    return SpecTypeId.ElectricalPotential;

                case ParameterTypeMap.ElectricalFrequency:
                    return SpecTypeId.ElectricalFrequency;
                //case ParameterTypeMap.ElectricalIlluminance:
                //     return SpecTypeId.;
                //case ParameterTypeMap.ElectricalLuminousFlux:
                //     return SpecTypeId.;
                case ParameterTypeMap.ElectricalPower:
                    return SpecTypeId.ElectricalPower;

                case ParameterTypeMap.HVACRoughness:
                    return SpecTypeId.HvacRoughness;
                //case ParameterTypeMap.ElectricalApparentPower:
                //     return SpecTypeId.;
                case ParameterTypeMap.ElectricalPowerDensity:
                    return SpecTypeId.ElectricalPowerDensity;

                case ParameterTypeMap.PipingDensity:
                    return SpecTypeId.PipingDensity;
                //case ParameterTypeMap.PipingFlow:
                //     return SpecTypeId.PipingFlow;
                case ParameterTypeMap.PipingFriction:
                    return SpecTypeId.PipingFriction;

                case ParameterTypeMap.PipingPressure:
                    return SpecTypeId.PipingPressure;

                case ParameterTypeMap.PipingTemperature:
                    return SpecTypeId.PipingTemperature;

                case ParameterTypeMap.PipingVelocity:
                    return SpecTypeId.PipingVelocity;

                case ParameterTypeMap.PipingViscosity:
                    return SpecTypeId.PipingViscosity;

                case ParameterTypeMap.PipeSize:
                    return SpecTypeId.PipeSize;

                case ParameterTypeMap.PipingRoughness:
                    return SpecTypeId.PipingRoughness;

                case ParameterTypeMap.Stress:
                    return SpecTypeId.Stress;

                case ParameterTypeMap.UnitWeight:
                    return SpecTypeId.UnitWeight;

                case ParameterTypeMap.ThermalExpansion:
                    return SpecTypeId.ThermalExpansionCoefficient;

                case ParameterTypeMap.LinearMoment:
                    return SpecTypeId.LinearMoment;
                //case ParameterTypeMap.ForcePerLength:
                //     return SpecTypeId.ForcePerLength;
                //case ParameterTypeMap.ForceLengthPerAngle:
                //     return SpecTypeId.ForceLengthPerAngle;
                //case ParameterTypeMap.LinearForcePerLength:
                //     return SpecTypeId.;
                //case ParameterTypeMap.LinearForceLengthPerAngle:
                //     return SpecTypeId.LinearForceLengthPerAngle;
                //case ParameterTypeMap.AreaForcePerLength:
                //     return SpecTypeId.;
                case ParameterTypeMap.PipingVolume:
                    return SpecTypeId.PipingVolume;

                case ParameterTypeMap.HVACViscosity:
                    return SpecTypeId.HvacViscosity;
                //case ParameterTypeMap.HVACCoefficientOfHeatTransfer:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACAirflowDensity:
                //     return SpecTypeId.;
                case ParameterTypeMap.Slope:
                    return SpecTypeId.Slope;
                //case ParameterTypeMap.HVACCoolingLoad:
                //     return SpecTypeId.HVACCoolingLoad;
                //case ParameterTypeMap.HVACCoolingLoadDividedByArea:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACCoolingLoadDividedByVolume:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACHeatingLoad:
                //     return SpecTypeId.HVACHeatingLoad;
                //case ParameterTypeMap.HVACHeatingLoadDividedByArea:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACHeatingLoadDividedByVolume:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACAirflowDividedByVolume:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACAirflowDividedByCoolingLoad:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACAreaDividedByCoolingLoad:
                //     return SpecTypeId.;
                //case ParameterTypeMap.WireSize:
                //     return SpecTypeId.;
                case ParameterTypeMap.HVACSlope:
                    return SpecTypeId.HvacSlope;

                case ParameterTypeMap.PipingSlope:
                    return SpecTypeId.PipingSlope;

                case ParameterTypeMap.Currency:
                    return SpecTypeId.Currency;
                //case ParameterTypeMap.ElectricalEfficacy:
                //     return SpecTypeId.;
                //case ParameterTypeMap.ElectricalWattage:
                //     return SpecTypeId.;
                case ParameterTypeMap.ColorTemperature:
                    return SpecTypeId.ColorTemperature;
                //case ParameterTypeMap.ElectricalLuminousIntensity:
                //     return SpecTypeId.ElectricalLuminousIntensity;
                //case ParameterTypeMap.ElectricalLuminance:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACAreaDividedByHeatingLoad:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACFactor:
                //     return SpecTypeId.;
                case ParameterTypeMap.ElectricalTemperature:
                    return SpecTypeId.ElectricalTemperature;
                //case ParameterTypeMap.ElectricalCableTraySize:
                //     return SpecTypeId.;
                //case ParameterTypeMap.ElectricalConduitSize:
                //     return SpecTypeId.;
                case ParameterTypeMap.ReinforcementVolume:
                    return SpecTypeId.ReinforcementVolume;

                case ParameterTypeMap.ReinforcementLength:
                    return SpecTypeId.ReinforcementLength;
                //case ParameterTypeMap.ElectricalDemandFactor:
                //     return SpecTypeId.ElectricalDemandFactor;
                //case ParameterTypeMap.HVACDuctInsulationThickness:
                //     return SpecTypeId.HVACDuctInsulationThickness;
                //case ParameterTypeMap.HVACDuctLiningThickness:
                //     return SpecTypeId.;
                case ParameterTypeMap.PipeInsulationThickness:
                    return SpecTypeId.PipeInsulationThickness;
                //case ParameterTypeMap.HVACThermalResistance:
                //     return SpecTypeId.;
                //case ParameterTypeMap.HVACThermalMass:
                //     return SpecTypeId.;
                case ParameterTypeMap.Acceleration:
                    return SpecTypeId.Acceleration;

                case ParameterTypeMap.BarDiameter:
                    return SpecTypeId.BarDiameter;

                case ParameterTypeMap.CrackWidth:
                    return SpecTypeId.CrackWidth;
                //case ParameterTypeMap.DisplacementDeflection:
                //     return SpecTypeId.DisplacementDeflection;
                case ParameterTypeMap.Energy:
                    return SpecTypeId.Energy;

                case ParameterTypeMap.StructuralFrequency:
                    return SpecTypeId.StructuralFrequency;

                case ParameterTypeMap.Mass:
                    return SpecTypeId.Mass;

                case ParameterTypeMap.MassPerUnitLength:
                    return SpecTypeId.MassPerUnitLength;

                case ParameterTypeMap.MomentOfInertia:
                    return SpecTypeId.MomentOfInertia;

                case ParameterTypeMap.SurfaceArea:
                    return SpecTypeId.SurfaceAreaPerUnitLength;

                case ParameterTypeMap.Period:
                    return SpecTypeId.Period;

                case ParameterTypeMap.Pulsation:
                    return SpecTypeId.Pulsation;

                case ParameterTypeMap.ReinforcementArea:
                    return SpecTypeId.ReinforcementArea;

                case ParameterTypeMap.ReinforcementAreaPerUnitLength:
                    return SpecTypeId.ReinforcementAreaPerUnitLength;

                case ParameterTypeMap.ReinforcementCover:
                    return SpecTypeId.ReinforcementCover;

                case ParameterTypeMap.ReinforcementSpacing:
                    return SpecTypeId.ReinforcementSpacing;

                case ParameterTypeMap.Rotation:
                    return SpecTypeId.Rotation;

                case ParameterTypeMap.SectionArea:
                    return SpecTypeId.SectionArea;

                case ParameterTypeMap.SectionDimension:
                    return SpecTypeId.SectionDimension;

                case ParameterTypeMap.SectionModulus:
                    return SpecTypeId.SectionModulus;

                case ParameterTypeMap.SectionProperty:
                    return SpecTypeId.SectionProperty;

                case ParameterTypeMap.StructuralVelocity:
                    return SpecTypeId.StructuralVelocity;

                case ParameterTypeMap.WarpingConstant:
                    return SpecTypeId.WarpingConstant;

                case ParameterTypeMap.Weight:
                    return SpecTypeId.Weight;

                case ParameterTypeMap.WeightPerUnitLength:
                    return SpecTypeId.WeightPerUnitLength;
                //case ParameterTypeMap.HVACThermalConductivity:
                //     return SpecTypeId.HVACThermalConductivity;
                //case ParameterTypeMap.HVACSpecificHeat:
                //     return SpecTypeId.HVACSpecificHeat;
                //case ParameterTypeMap.HVACSpecificHeatOfVaporization:
                //    return SpecTypeId.HVACSpecificHeatOfVaporization;
                //case ParameterTypeMap.HVACPermeability:
                //     return SpecTypeId.;
                case ParameterTypeMap.ElectricalResistivity:
                    return SpecTypeId.ElectricalResistivity;

                case ParameterTypeMap.MassDensity:
                    return SpecTypeId.MassDensity;

                case ParameterTypeMap.MassPerUnitArea:
                    return SpecTypeId.MassPerUnitArea;

                case ParameterTypeMap.PipeDimension:
                    return SpecTypeId.PipeDimension;
                //case ParameterTypeMap.PipeMass:
                //     return SpecTypeId.PipeMassPerUnitLength;
                case ParameterTypeMap.PipeMassPerUnitLength:
                    return SpecTypeId.PipeMassPerUnitLength;

                case ParameterTypeMap.HVACTemperatureDifference:
                    return SpecTypeId.HvacTemperatureDifference;

                case ParameterTypeMap.PipingTemperatureDifference:
                    return SpecTypeId.PipingTemperatureDifference;

                case ParameterTypeMap.ElectricalTemperatureDifference:
                    return SpecTypeId.ElectricalTemperatureDifference;

                default:
                    return SpecTypeId.String.Text;
            }
        }

#endif
    }
}