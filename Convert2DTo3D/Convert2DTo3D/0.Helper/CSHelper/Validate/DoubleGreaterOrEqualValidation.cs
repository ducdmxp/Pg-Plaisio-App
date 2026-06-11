using Convert2DTo3D;
using Convert2DTo3D.Properties;
using System;
using System.Globalization;
using System.Windows.Controls;

namespace Kajima
{
    public class DoubleGreaterOrEqualValidation : System.Windows.Controls.ValidationRule
    {
        private double _Min = 0;

        public double Min
        {
            get { return _Min; }
            set { _Min = value; }
        }

        private string message = "";

        public DoubleGreaterOrEqualValidation()
        { }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            message = MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_InputRequireDoubleAndGreaterEqualValue);
            message = string.Format(message, Min.ToString());
            try
            {
                string input = (string)value;
                if (input.EndsWith(".") || input == "-0" || (input.Length == 1 && input.StartsWith("-")) || (input.Contains(".") && input.EndsWith("0"))) return new ValidationResult(false, message);

                double output = 0;
                bool check = double.TryParse((string)value, out output);
                if (check && (output > Min || NumberUtils.IsEquals(output, Min)) && string.IsNullOrEmpty((string)value) == false)
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