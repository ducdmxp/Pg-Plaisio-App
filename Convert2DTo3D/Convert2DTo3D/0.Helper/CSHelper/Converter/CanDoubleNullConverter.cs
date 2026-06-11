using System;
using System.Globalization;
using System.Windows.Data;

namespace Convert2DTo3D
{
    public class CanDoubleNullConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string strValue = value as string;
            if (string.IsNullOrWhiteSpace(strValue)) return null;
            double output = 0;
            if (double.TryParse(strValue, out output) == false) return null;
            return value;
        }
    }
}