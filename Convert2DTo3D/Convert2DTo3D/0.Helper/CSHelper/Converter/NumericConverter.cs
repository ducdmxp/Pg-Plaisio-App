using System;
using System.Globalization;
using System.Windows.Data;

namespace Convert2DTo3D
{
    public class NumericConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() ?? "0";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (double.TryParse(value?.ToString(), out double result))
            {
                return result;
            }
            return 0;
        }
    }
}