using Convert2DTo3D;
using Convert2DTo3D.Properties;
using System;
using System.Globalization;
using System.Windows.Controls;

namespace Kajima
{
    public class IntegerValidation : System.Windows.Controls.ValidationRule
    {
        private string message = "";

        public IntegerValidation()
        {
            message = MngLanguage.GetNameFromCurrentResource(() => Resources_en_US.GEN_InputRequireDouble);
        }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            try
            {
                string input = (string)value;
                if (input.EndsWith(".") || (input.Length == 1 && input.StartsWith("-")) || (input.Contains(".") && input.EndsWith("0"))) return new ValidationResult(false, message);
                int output = 0;
                bool check = int.TryParse((string)value, out output);
                if (check && string.IsNullOrEmpty((string)value) == false)
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