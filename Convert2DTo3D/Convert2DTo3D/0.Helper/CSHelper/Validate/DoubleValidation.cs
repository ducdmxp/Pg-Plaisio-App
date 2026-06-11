using Convert2DTo3D;
using Convert2DTo3D.Properties;
using System;
using System.Globalization;
using System.Windows.Controls;

namespace Kajima
{
    public class DoubleValidation : System.Windows.Controls.ValidationRule
    {
        private string message = "";

        public DoubleValidation()
        {
            message = MngLanguage.GetNameFromCurrentResource(() => Resources_en_US.GEN_InputRequireInteger);
        }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            try
            {
                string input = (string)value;
                if (input.EndsWith(".") || input == "-0" || (input.Length == 1 && input.StartsWith("-")) || (input.Contains(".") && input.EndsWith("0"))) return new ValidationResult(false, message);
                double output = 0;
                bool check = double.TryParse((string)value, out output);
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