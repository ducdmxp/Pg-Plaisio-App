using Convert2DTo3D;
using Convert2DTo3D.Properties;
using System;
using System.Globalization;
using System.Windows.Controls;

namespace Kajima
{
    public class RequireValidation : System.Windows.Controls.ValidationRule
    {
        private string message = "";

        public RequireValidation()
        {
            message = MngLanguage.GetNameFromCurrentResource(() => Resources_en_US.GEN_InputRequireValidation);
        }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            try
            {
                if (string.IsNullOrEmpty(((string)value).Trim()) == false)
                {
                    return ValidationResult.ValidResult;
                }
                else
                {
                    return new ValidationResult(false, message);
                }
            }
            catch (Exception ex)
            {
                return new ValidationResult(false, ex.Message);
            }
        }
    }
}