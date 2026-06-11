using System.Diagnostics;
using System.Text;
using System.Windows.Controls;

namespace Convert2DTo3D
{
    public class TraceValidationRule : ValidationRule
    {
        public string PropertyName
        {
            get;
            set;
        }

        public override ValidationResult Validate(object value, System.Globalization.CultureInfo cultureInfo)
        {
            StringBuilder buidler = new StringBuilder();
            Debug.WriteLine(buidler.Append("TraceValidationRule for '")
                                .Append(PropertyName)
                                .Append("' called. ValidationStep='")
                                .Append(ValidationStep.ToString())
                                .Append("'").ToString());

            return ValidationResult.ValidResult;  // Don't stop the validation process by reporting an error.
        }
    }
}