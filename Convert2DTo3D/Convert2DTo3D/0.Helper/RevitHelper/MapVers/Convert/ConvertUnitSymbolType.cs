using Autodesk.Revit.DB;
using System;

namespace Convert2DTo3D
{
    /*
	List<string> forceTypeNames = new List<string>();
	Type unitTypeIdType = typeof(SymbolTypeId);
	PropertyInfo[] properties = unitTypeIdType.GetProperties(BindingFlags.Static | BindingFlags.Public);
	foreach (PropertyInfo property in properties) forceTypeNames.Add(property.Name);
	foreach (int i in Enum.GetValues(typeof(UnitSymblTypeMap)))
	{
		var enumValuen = Enum.GetName(typeof(UnitSymblTypeMap), i);
		enumValuen = enumValuen.Replace("UST", "").ToLower();
		var enumValue = enumValuen.Split('_');

		var name = string.Join("", enumValue);
		var chek = forceTypeNames.Where(k => k.ToLower().Contains(name) || k.ToLower() == name).ToList();

		if (chek.Count() == 0)
		{
			Debug.WriteLine($"case UnitSymblTypeMap.{Enum.GetName(typeof(UnitSymblTypeMap), i)}: \n\treturn SymbolTypeId.Valid;");
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
			Debug.WriteLine($"case UnitSymblTypeMap.{Enum.GetName(typeof(UnitSymblTypeMap), i)}: \n\t  return SymbolTypeId.{ne};");
		}
	}
	 */

    public class ConvertUnitSymbolType
    {
#if REVIT2019 || REVIT2020 || REVIT2021

        public static UnitSymbolType Convert(UnitSymbolTypeMap unitSymblTypeMap)
        {
            var stringEnum = unitSymblTypeMap.ToString();
            var checkExists = Enum.IsDefined(typeof(UnitSymbolType), stringEnum);
            if (checkExists == true)
            {
                UnitSymbolType output = UnitSymbolType.UST_M;
                Enum.TryParse<UnitSymbolType>(stringEnum, out output);
                return output;
            }
            return UnitSymbolType.UST_M;
        }

#else

        public static ForgeTypeId Convert(UnitSymbolTypeMap unitSymblTypeMap)
        {
            switch (unitSymblTypeMap)
            {
                case UnitSymbolTypeMap.UST_M:
                    return SymbolTypeId.Meter; //
                case UnitSymbolTypeMap.UST_CM:
                    return SymbolTypeId.Cm;

                case UnitSymbolTypeMap.UST_MM:
                    return SymbolTypeId.Mm;

                case UnitSymbolTypeMap.UST_LF:
                    return SymbolTypeId.Lf;

                case UnitSymbolTypeMap.UST_FOOT_SINGLE_QUOTE:
                    return SymbolTypeId.FootSingleQuote;

                case UnitSymbolTypeMap.UST_FT:
                    return SymbolTypeId.Ft;

                case UnitSymbolTypeMap.UST_INCH_DOUBLE_QUOTE:
                    return SymbolTypeId.InchDoubleQuote;

                case UnitSymbolTypeMap.UST_IN:
                    return SymbolTypeId.In;

                case UnitSymbolTypeMap.UST_ACRES:
                    return SymbolTypeId.Acres;

                case UnitSymbolTypeMap.UST_HECTARES:
                    return SymbolTypeId.Hectare;

                case UnitSymbolTypeMap.UST_CY:
                    return SymbolTypeId.Cy;

                case UnitSymbolTypeMap.UST_SF:
                    return SymbolTypeId.Sf;

                case UnitSymbolTypeMap.UST_FT_SUP_2:
                    return SymbolTypeId.FtSup2;

                case UnitSymbolTypeMap.UST_FT_CARET_2:
                    return SymbolTypeId.FtCaret2;

                case UnitSymbolTypeMap.UST_M_SUP_2:
                    return SymbolTypeId.MSup2;

                case UnitSymbolTypeMap.UST_M_CARET_2:
                    return SymbolTypeId.MCaret2;

                case UnitSymbolTypeMap.UST_CF:
                    return SymbolTypeId.Cf;

                case UnitSymbolTypeMap.UST_FT_SUP_3:
                    return SymbolTypeId.FtSup3;

                case UnitSymbolTypeMap.UST_FT_CARET_3:
                    return SymbolTypeId.FtCaret3;

                case UnitSymbolTypeMap.UST_M_SUP_3:
                    return SymbolTypeId.MSup3;

                case UnitSymbolTypeMap.UST_M_CARET_3:
                    return SymbolTypeId.MCaret3;

                case UnitSymbolTypeMap.UST_DEGREE_SYMBOL:
                    return SymbolTypeId.Degree;

                case UnitSymbolTypeMap.UST_PERCENT_SIGN:
                    return SymbolTypeId.Percent;

                case UnitSymbolTypeMap.UST_IN_SUP_2:
                    return SymbolTypeId.InSup2;

                case UnitSymbolTypeMap.UST_IN_CARET_2:
                    return SymbolTypeId.InCaret2;

                case UnitSymbolTypeMap.UST_CM_SUP_2:
                    return SymbolTypeId.CmSup2;

                case UnitSymbolTypeMap.UST_CM_CARET_2:
                    return SymbolTypeId.CmCaret2;

                case UnitSymbolTypeMap.UST_MM_SUP_2:
                    return SymbolTypeId.MmSup2;

                case UnitSymbolTypeMap.UST_MM_CARET_2:
                    return SymbolTypeId.MmCaret2;

                case UnitSymbolTypeMap.UST_IN_SUP_3:
                    return SymbolTypeId.InSup3;

                case UnitSymbolTypeMap.UST_IN_CARET_3:
                    return SymbolTypeId.InCaret3;

                case UnitSymbolTypeMap.UST_CM_SUP_3:
                    return SymbolTypeId.CmSup3;

                case UnitSymbolTypeMap.UST_CM_CARET_3:
                    return SymbolTypeId.CmCaret3;

                case UnitSymbolTypeMap.UST_MM_SUP_3:
                    return SymbolTypeId.MmSup3;

                case UnitSymbolTypeMap.UST_MM_CARET_3:
                    return SymbolTypeId.MmCaret3;

                case UnitSymbolTypeMap.UST_L:
                    return SymbolTypeId.Liter; //
                case UnitSymbolTypeMap.UST_GAL:
                    return SymbolTypeId.Gal;

                case UnitSymbolTypeMap.UST_KG_PER_CU_M:
                    return SymbolTypeId.KgPerM; //
                case UnitSymbolTypeMap.UST_LB_MASS_PER_CU_FT:
                    return SymbolTypeId.LbfPerFt; //
                case UnitSymbolTypeMap.UST_LBM_PER_CU_FT:
                    return SymbolTypeId.LbmPerFt; //
                case UnitSymbolTypeMap.UST_BTU:
                    return SymbolTypeId.Btu;

                case UnitSymbolTypeMap.UST_CAL:
                    return SymbolTypeId.Cal;

                case UnitSymbolTypeMap.UST_KCAL:
                    return SymbolTypeId.Kcal;

                case UnitSymbolTypeMap.UST_JOULE:
                    return SymbolTypeId.Joule;

                case UnitSymbolTypeMap.UST_KWH:
                    return SymbolTypeId.KWh;

                case UnitSymbolTypeMap.UST_THERM:
                    return SymbolTypeId.Therm;

                case UnitSymbolTypeMap.UST_PASCAL_PER_M:
                    return SymbolTypeId.PaPerM;

                case UnitSymbolTypeMap.UST_WATT:
                    return SymbolTypeId.Watt;

                case UnitSymbolTypeMap.UST_KILOWATT:
                    return SymbolTypeId.KW;

                case UnitSymbolTypeMap.UST_BTU_PER_S:
                    return SymbolTypeId.BtuPerS;

                case UnitSymbolTypeMap.UST_BTU_PER_H:
                    return SymbolTypeId.BtuPerH;

                case UnitSymbolTypeMap.UST_CAL_PER_S:
                    return SymbolTypeId.CalPerS;

                case UnitSymbolTypeMap.UST_KCAL_PER_S:
                    return SymbolTypeId.KcalPerS;

                case UnitSymbolTypeMap.UST_PASCAL:
                    return SymbolTypeId.Pa;

                case UnitSymbolTypeMap.UST_KILOPASCAL:
                    return SymbolTypeId.KPa;

                case UnitSymbolTypeMap.UST_MEGAPASCAL:
                    return SymbolTypeId.MPa;

                case UnitSymbolTypeMap.UST_PSI:
                    return SymbolTypeId.Psi;

                case UnitSymbolTypeMap.UST_PSIG:
                    return SymbolTypeId.Psig;

                case UnitSymbolTypeMap.UST_PSIA:
                    return SymbolTypeId.Psia;

                case UnitSymbolTypeMap.UST_IN_HG:
                    return SymbolTypeId.InHg;

                case UnitSymbolTypeMap.UST_MM_HG:
                    return SymbolTypeId.MmHg;

                case UnitSymbolTypeMap.UST_ATM:
                    return SymbolTypeId.Atm;

                case UnitSymbolTypeMap.UST_BAR:
                    return SymbolTypeId.Bar;

                case UnitSymbolTypeMap.UST_DEGREE_F:
                    return SymbolTypeId.DegreeF;

                case UnitSymbolTypeMap.UST_DEGREE_C:
                    return SymbolTypeId.DegreeC;

                case UnitSymbolTypeMap.UST_KELVIN:
                    return SymbolTypeId.Kelvin;

                case UnitSymbolTypeMap.UST_DEGREE_R:
                    return SymbolTypeId.DegreeR;

                case UnitSymbolTypeMap.UST_FT_PER_MIN:
                    return SymbolTypeId.FtPerMin;

                case UnitSymbolTypeMap.UST_FPM:
                    return SymbolTypeId.Fpm;

                case UnitSymbolTypeMap.UST_M_PER_S:
                    return SymbolTypeId.MPerS;

                case UnitSymbolTypeMap.UST_CM_PER_MIN:
                    return SymbolTypeId.CmPerMin;

                case UnitSymbolTypeMap.UST_CFM:
                    return SymbolTypeId.Cfm;

                case UnitSymbolTypeMap.UST_L_PER_S:
                    return SymbolTypeId.LPerS;

                case UnitSymbolTypeMap.UST_LPS:
                    return SymbolTypeId.Lps;

                case UnitSymbolTypeMap.UST_CMS:
                    return SymbolTypeId.Cms;

                case UnitSymbolTypeMap.UST_CMH:
                    return SymbolTypeId.Cmh;

                case UnitSymbolTypeMap.UST_GAL_PER_MIN:
                    return SymbolTypeId.GalPerMin;

                case UnitSymbolTypeMap.UST_GPM:
                    return SymbolTypeId.Gpm;

                case UnitSymbolTypeMap.UST_USGPM:
                    return SymbolTypeId.Usgpm;

                case UnitSymbolTypeMap.UST_GAL_PER_H:
                    return SymbolTypeId.GalPerH;

                case UnitSymbolTypeMap.UST_GPH:
                    return SymbolTypeId.Gph;

                case UnitSymbolTypeMap.UST_USGPH:
                    return SymbolTypeId.Usgph;

                case UnitSymbolTypeMap.UST_AMPERE:
                    return SymbolTypeId.Ampere;

                case UnitSymbolTypeMap.UST_KILOAMPERE:
                    return SymbolTypeId.KA;

                case UnitSymbolTypeMap.UST_MILLIAMPERE:
                    return SymbolTypeId.MA;

                case UnitSymbolTypeMap.UST_VOLT:
                    return SymbolTypeId.Volt;

                case UnitSymbolTypeMap.UST_KILOVOLT:
                    return SymbolTypeId.KV;

                case UnitSymbolTypeMap.UST_MILLIVOLT:
                    return SymbolTypeId.MV;

                case UnitSymbolTypeMap.UST_HZ:
                    return SymbolTypeId.Hz;

                case UnitSymbolTypeMap.UST_CPS:
                    return SymbolTypeId.Cps;

                case UnitSymbolTypeMap.UST_LX:
                    return SymbolTypeId.Lx;

                case UnitSymbolTypeMap.UST_FC:
                    return SymbolTypeId.Fc;

                case UnitSymbolTypeMap.UST_FTC:
                    return SymbolTypeId.Ftc;

                case UnitSymbolTypeMap.UST_FL:
                    return SymbolTypeId.FL;

                case UnitSymbolTypeMap.UST_FL_LOWERCASE:
                    return SymbolTypeId.FlLowercase;

                case UnitSymbolTypeMap.UST_FTL:
                    return SymbolTypeId.FtL;

                case UnitSymbolTypeMap.UST_CD:
                    return SymbolTypeId.Cd;

                case UnitSymbolTypeMap.UST_LM:
                    return SymbolTypeId.Lm;

                case UnitSymbolTypeMap.UST_VOLTAMPERE:
                    return SymbolTypeId.VA;

                case UnitSymbolTypeMap.UST_KILOVOLTAMPERE:
                    return SymbolTypeId.KVA;

                case UnitSymbolTypeMap.UST_HP:
                    return SymbolTypeId.Hp;

                case UnitSymbolTypeMap.UST_N:
                    return SymbolTypeId.Newton;

                case UnitSymbolTypeMap.UST_DA_N:
                    return SymbolTypeId.DaN;

                case UnitSymbolTypeMap.UST_K_N:
                    return SymbolTypeId.KN;

                case UnitSymbolTypeMap.UST_M_N:
                    return SymbolTypeId.MN;

                case UnitSymbolTypeMap.UST_KIP:
                    return SymbolTypeId.Kip;

                case UnitSymbolTypeMap.UST_KGF:
                    return SymbolTypeId.Kgf;

                case UnitSymbolTypeMap.UST_TF:
                    return SymbolTypeId.Tf;

                case UnitSymbolTypeMap.UST_LB_FORCE:
                    return SymbolTypeId.LbForce;

                case UnitSymbolTypeMap.UST_LBF:
                    return SymbolTypeId.Lbf;

                case UnitSymbolTypeMap.UST_N_PER_M:
                    return SymbolTypeId.NPerM;

                case UnitSymbolTypeMap.UST_DA_N_PER_M:
                    return SymbolTypeId.DaNPerM;

                case UnitSymbolTypeMap.UST_K_N_PER_M:
                    return SymbolTypeId.KNPerM;

                case UnitSymbolTypeMap.UST_M_N_PER_M:
                    return SymbolTypeId.MNPerM;

                case UnitSymbolTypeMap.UST_KIP_PER_FT:
                    return SymbolTypeId.KipPerFt;

                case UnitSymbolTypeMap.UST_KGF_PER_M:
                    return SymbolTypeId.KgfPerM;

                case UnitSymbolTypeMap.UST_TF_PER_M:
                    return SymbolTypeId.TfPerM;

                case UnitSymbolTypeMap.UST_LB_FORCE_PER_FT:
                    return SymbolTypeId.LbForcePerFt;

                case UnitSymbolTypeMap.UST_LBF_PER_FT:
                    return SymbolTypeId.LbfPerFt;

                case UnitSymbolTypeMap.UST_N_PER_M_SUP_2:
                    return SymbolTypeId.NPerMSup2;

                case UnitSymbolTypeMap.UST_DA_N_PER_M_SUP_2:
                    return SymbolTypeId.DaNPerMSup2;

                case UnitSymbolTypeMap.UST_K_N_PER_M_SUP_2:
                    return SymbolTypeId.KNPerMSup2;

                case UnitSymbolTypeMap.UST_M_N_PER_M_SUP_2:
                    return SymbolTypeId.MNPerMSup2;

                case UnitSymbolTypeMap.UST_KSF:
                    return SymbolTypeId.Ksf;

                case UnitSymbolTypeMap.UST_KGF_PER_M_SUP_2:
                    return SymbolTypeId.KgfPerMSup2;

                case UnitSymbolTypeMap.UST_TF_PER_M_SUP_2:
                    return SymbolTypeId.TfPerMSup2;

                case UnitSymbolTypeMap.UST_PSF:
                    return SymbolTypeId.Psf;

                case UnitSymbolTypeMap.UST_N_DASH_M:
                    return SymbolTypeId.NDashM;

                case UnitSymbolTypeMap.UST_DA_N_DASH_M:
                    return SymbolTypeId.DaNDashM;

                case UnitSymbolTypeMap.UST_K_N_DASH_M:
                    return SymbolTypeId.KNDashM;

                case UnitSymbolTypeMap.UST_M_N_DASH_M:
                    return SymbolTypeId.MNDashM;

                case UnitSymbolTypeMap.UST_KIP_DASH_FT:
                    return SymbolTypeId.KipDashFt;

                case UnitSymbolTypeMap.UST_KGF_DASH_M:
                    return SymbolTypeId.KgfDashM;

                case UnitSymbolTypeMap.UST_TF_DASH_M:
                    return SymbolTypeId.TfDashM;

                case UnitSymbolTypeMap.UST_LB_FORCE_DASH_FT:
                    return SymbolTypeId.LbForceDashFt;

                case UnitSymbolTypeMap.UST_LBF_DASH_FT:
                    return SymbolTypeId.LbfDashFt;

                case UnitSymbolTypeMap.UST_M_PER_K_N:
                    return SymbolTypeId.MPerKN;

                case UnitSymbolTypeMap.UST_FT_PER_KIP:
                    return SymbolTypeId.FtPerKip;

                case UnitSymbolTypeMap.UST_M_SUP_2_PER_K_N:
                    return SymbolTypeId.MSup2PerKN;

                case UnitSymbolTypeMap.UST_FT_SUP_2_PER_KIP:
                    return SymbolTypeId.FtSup2PerKip;

                case UnitSymbolTypeMap.UST_M_SUP_3_PER_K_N:
                    return SymbolTypeId.MSup3PerKN;

                case UnitSymbolTypeMap.UST_FT_SUP_3_PER_KIP:
                    return SymbolTypeId.FtSup3PerKip;

                case UnitSymbolTypeMap.UST_INV_K_N:
                    return SymbolTypeId.InvKN;

                case UnitSymbolTypeMap.UST_INV_KIP:
                    return SymbolTypeId.InvKip;

                case UnitSymbolTypeMap.UST_FTH2O_PER_100FT:
                    return SymbolTypeId.FtH2OPer100ft;

                case UnitSymbolTypeMap.UST_FT_OF_WATER_PER_100FT:
                    return SymbolTypeId.FtOfWaterPer100ft;

                case UnitSymbolTypeMap.UST_FEET_OF_WATER_PER_100FT:
                    return SymbolTypeId.FeetOfWaterPer100ft;

                case UnitSymbolTypeMap.UST_FTH2O:
                    return SymbolTypeId.FtH2O;

                case UnitSymbolTypeMap.UST_FT_OF_WATER:
                    return SymbolTypeId.FtOfWater;

                case UnitSymbolTypeMap.UST_FEET_OF_WATER:
                    return SymbolTypeId.FeetOfWater;

                case UnitSymbolTypeMap.UST_PA_S:
                    return SymbolTypeId.NgPerPaSMSup2;

                case UnitSymbolTypeMap.UST_LBM_PER_FT_S:
                    return SymbolTypeId.LbmPerFtSup3;

                case UnitSymbolTypeMap.UST_CP:
                    return SymbolTypeId.CP;

                case UnitSymbolTypeMap.UST_FT_PER_S:
                    return SymbolTypeId.FtPerS;

                case UnitSymbolTypeMap.UST_FPS:
                    return SymbolTypeId.Fps;

                case UnitSymbolTypeMap.UST_KSI:
                    return SymbolTypeId.Ksi;

                case UnitSymbolTypeMap.UST_KN_PER_M_SUP_3:
                    return SymbolTypeId.KNPerMSup3;

                case UnitSymbolTypeMap.UST_KIP_PER_IN_SUP_3:
                    return SymbolTypeId.KipPerInSup3;

                case UnitSymbolTypeMap.UST_INV_DEGREE_F:
                    return SymbolTypeId.InvDegreeF;

                case UnitSymbolTypeMap.UST_INV_DEGREE_C:
                    return SymbolTypeId.InvDegreeC;

                case UnitSymbolTypeMap.UST_N_DASH_M_PER_M:
                    return SymbolTypeId.NDashMPerM;

                case UnitSymbolTypeMap.UST_DA_N_DASH_M_PER_M:
                    return SymbolTypeId.DaNDashMPerM;

                case UnitSymbolTypeMap.UST_K_N_DASH_M_PER_M:
                    return SymbolTypeId.KNDashMPerM;

                case UnitSymbolTypeMap.UST_M_N_DASH_M_PER_M:
                    return SymbolTypeId.MNDashMPerM;

                case UnitSymbolTypeMap.UST_KIP_DASH_FT_PER_FT:
                    return SymbolTypeId.KipDashFtPerFt;

                case UnitSymbolTypeMap.UST_KGF_DASH_M_PER_M:
                    return SymbolTypeId.KgfDashMPerM;

                case UnitSymbolTypeMap.UST_TF_DASH_M_PER_M:
                    return SymbolTypeId.TfDashMPerM;

                case UnitSymbolTypeMap.UST_LB_FORCE_DASH_FT_PER_FT:
                    return SymbolTypeId.LbForceDashFtPerFt;

                case UnitSymbolTypeMap.UST_LBF_DASH_FT_PER_FT:
                    return SymbolTypeId.LbfDashFtPerFt;

                case UnitSymbolTypeMap.UST_COLON_10:
                    return SymbolTypeId.Colon10;

                case UnitSymbolTypeMap.UST_COLON_12:
                    return SymbolTypeId.Colon12;

                case UnitSymbolTypeMap.UST_TON:
                    return SymbolTypeId.Ton;

                case UnitSymbolTypeMap.UST_TON_OF_REFRIGERATION:
                    return SymbolTypeId.TonOfRefrigeration;

                case UnitSymbolTypeMap.UST_CFM_PER_CF:
                    return SymbolTypeId.CfmPerCf;

                case UnitSymbolTypeMap.UST_CFM_PER_TON:
                    return SymbolTypeId.CfmPerTon;

                case UnitSymbolTypeMap.UST_L_PER_S_KW:
                    return SymbolTypeId.LPerSKw;

                case UnitSymbolTypeMap.UST_SF_PER_TON:
                    return SymbolTypeId.SfPerTon;

                case UnitSymbolTypeMap.UST_DOLLAR:
                    return SymbolTypeId.UsDollar;

                case UnitSymbolTypeMap.UST_EURO_SUFFIX:
                    return SymbolTypeId.EuroSuffix;

                case UnitSymbolTypeMap.UST_EURO_PREFIX:
                    return SymbolTypeId.EuroPrefix;

                case UnitSymbolTypeMap.UST_POUND:
                    return SymbolTypeId.UkPound;

                case UnitSymbolTypeMap.UST_YEN:
                    return SymbolTypeId.Yen;

                case UnitSymbolTypeMap.UST_WON:
                    return SymbolTypeId.Won;

                case UnitSymbolTypeMap.UST_DONG:
                    return SymbolTypeId.Dong;

                case UnitSymbolTypeMap.UST_BAHT:
                    return SymbolTypeId.Baht;

                case UnitSymbolTypeMap.UST_LM_PER_W:
                    return SymbolTypeId.LmPerW;

                case UnitSymbolTypeMap.UST_SF_PER_MBH:
                    return SymbolTypeId.SfPerMbh;

                case UnitSymbolTypeMap.UST_K_N_PER_CM_SUP_2:
                    return SymbolTypeId.KNPerCmSup2;

                case UnitSymbolTypeMap.UST_N_PER_MM_SUP_2:
                    return SymbolTypeId.NPerMmSup2;

                case UnitSymbolTypeMap.UST_K_N_PER_MM_SUP_2:
                    return SymbolTypeId.KNPerMmSup2;

                case UnitSymbolTypeMap.UST_ONE_COLON:
                    return SymbolTypeId.OneColon;

                case UnitSymbolTypeMap.UST_BTU_PER_F:
                    return SymbolTypeId.BtuPerFtSup2DegreeF;

                case UnitSymbolTypeMap.UST_TM:
                    return SymbolTypeId.Atm;

                case UnitSymbolTypeMap.UST_LB_MASS:
                    return SymbolTypeId.LbMass;

                case UnitSymbolTypeMap.UST_LBM:
                    return SymbolTypeId.Lbm;

                case UnitSymbolTypeMap.UST_FT_SUP_4:
                    return SymbolTypeId.FtSup4;

                case UnitSymbolTypeMap.UST_IN_SUP_4:
                    return SymbolTypeId.InSup4;

                case UnitSymbolTypeMap.UST_MM_SUP_4:
                    return SymbolTypeId.MmSup4;

                case UnitSymbolTypeMap.UST_CM_SUP_4:
                    return SymbolTypeId.CmSup4;

                case UnitSymbolTypeMap.UST_M_SUP_4:
                    return SymbolTypeId.MSup4;

                case UnitSymbolTypeMap.UST_FT_SUP_6:
                    return SymbolTypeId.FtSup6;

                case UnitSymbolTypeMap.UST_IN_SUP_6:
                    return SymbolTypeId.InSup6;

                case UnitSymbolTypeMap.UST_MM_SUP_6:
                    return SymbolTypeId.MmSup6;

                case UnitSymbolTypeMap.UST_CM_SUP_6:
                    return SymbolTypeId.CmSup6;

                case UnitSymbolTypeMap.UST_M_SUP_6:
                    return SymbolTypeId.MSup6;

                case UnitSymbolTypeMap.UST_LB_MASS_PER_FT:
                    return SymbolTypeId.LbMassPerFt;

                case UnitSymbolTypeMap.UST_LBM_PER_FT:
                    return SymbolTypeId.LbmPerFt;

                case UnitSymbolTypeMap.UST_RAD:
                    return SymbolTypeId.Rad;

                case UnitSymbolTypeMap.UST_GRAD:
                    return SymbolTypeId.Grad;

                case UnitSymbolTypeMap.UST_RAD_PER_S:
                    return SymbolTypeId.RadPerS;

                case UnitSymbolTypeMap.UST_MS:
                    return SymbolTypeId.Ms;

                case UnitSymbolTypeMap.UST_S:
                    return SymbolTypeId.Second; //
                case UnitSymbolTypeMap.UST_MIN:
                    return SymbolTypeId.Min;

                case UnitSymbolTypeMap.UST_H:
                    return SymbolTypeId.Hour;

                case UnitSymbolTypeMap.UST_KM_PER_H:
                    return SymbolTypeId.KmPerH;

                case UnitSymbolTypeMap.UST_KJ:
                    return SymbolTypeId.KJ;

                case UnitSymbolTypeMap.UST_J_PER_G_CELSIUS:
                    return SymbolTypeId.JPerGDegreeC;

                case UnitSymbolTypeMap.UST_J_PER_G:
                    return SymbolTypeId.JPerG;

                case UnitSymbolTypeMap.UST_NG_PER_PA_S_SQ_M:
                    return SymbolTypeId.NgPerPaSMSup2;

                case UnitSymbolTypeMap.UST_OHM_M:
                    return SymbolTypeId.OhmM;

                case UnitSymbolTypeMap.UST_BTU_PER_H_FT_DEGREE_F:
                    return SymbolTypeId.BtuPerHFtDegreeF;

                case UnitSymbolTypeMap.UST_BTU_PER_LB_DEGREE_F:
                    return SymbolTypeId.BtuPerLbDegreeF;

                case UnitSymbolTypeMap.UST_BTU_PER_LB:
                    return SymbolTypeId.BtuPerLb;

                case UnitSymbolTypeMap.UST_GR_PER_H_SQ_FT_IN_HG:
                    return SymbolTypeId.GrPerHFtSup2InHg;

                case UnitSymbolTypeMap.UST_PER_MILLE_SIGN:
                    return SymbolTypeId.PerMille;

                case UnitSymbolTypeMap.UST_DM:
                    return SymbolTypeId.Dm;

                case UnitSymbolTypeMap.UST_J_PER_KG_CELSIUS:
                    return SymbolTypeId.JPerKgDegreeC;

                case UnitSymbolTypeMap.UST_UM_PER_M_C:
                    return SymbolTypeId.UmPerMDegreeC;

                case UnitSymbolTypeMap.UST_UIN_PER_IN_F:
                    return SymbolTypeId.UinPerInDegreeF; //
                case UnitSymbolTypeMap.UST_USTONNES_MASS_TONS:
                    return SymbolTypeId.UsTonnesMassTons;

                case UnitSymbolTypeMap.UST_USTONNES_MASS_T:
                    return SymbolTypeId.UsTonnesMassT;

                case UnitSymbolTypeMap.UST_USTONNES_MASS_ST:
                    return SymbolTypeId.UsTonnesMassSt;

                case UnitSymbolTypeMap.UST_USTONNES_FORCE_TONSF:
                    return SymbolTypeId.UsTonnesForceTons;

                case UnitSymbolTypeMap.UST_USTONNES_FORCE_STF:
                    return SymbolTypeId.UsTonnesForceSt; //
                case UnitSymbolTypeMap.UST_USTONNES_FORCE_AS_MASS_TONS:
                    return SymbolTypeId.UsTonnesMassTons;

                case UnitSymbolTypeMap.UST_USTONNES_FORCE_AS_MASS_T:
                    return SymbolTypeId.UsTonnesForceT;

                case UnitSymbolTypeMap.UST_USTONNES_FORCE_AS_MASS_ST:
                    return SymbolTypeId.UsTonnesForceSt;

                case UnitSymbolTypeMap.UST_L_PER_M:
                    return SymbolTypeId.LPerMin;

                case UnitSymbolTypeMap.UST_LPM:
                    return SymbolTypeId.Lpm;

                case UnitSymbolTypeMap.UST_DEGREE_F_DIFFERENCE:
                    return SymbolTypeId.DegreeFInterval;

                case UnitSymbolTypeMap.UST_DELTA_DEGREE_F:
                    return SymbolTypeId.DeltaDegreeF;

                case UnitSymbolTypeMap.UST_DEGREE_C_DIFFERENCE:
                    return SymbolTypeId.DegreeCInterval;

                case UnitSymbolTypeMap.UST_DELTA_DEGREE_C:
                    return SymbolTypeId.DeltaDegreeC;

                case UnitSymbolTypeMap.UST_KELVIN_DIFFERENCE:
                    return SymbolTypeId.KelvinInterval;

                case UnitSymbolTypeMap.UST_DELTA_KELVIN:
                    return SymbolTypeId.DeltaDegreeR;

                case UnitSymbolTypeMap.UST_DEGREE_R_DIFFERENCE:
                    return SymbolTypeId.DegreeRInterval;

                case UnitSymbolTypeMap.UST_DELTA_DEGREE_R:
                    return SymbolTypeId.DeltaDegreeR;

                default:
                    return SymbolTypeId.DeltaDegreeR;
                    break;
            }
        }

#endif
    }
}