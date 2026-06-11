using System;
using System.Globalization;
using System.Windows.Data;

namespace Convert2DTo3D
{
    public class Bool2VisibilityCollapsedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                bool boolValue = (bool)value;
                return boolValue == true ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            }
            catch (Exception)
            { }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}