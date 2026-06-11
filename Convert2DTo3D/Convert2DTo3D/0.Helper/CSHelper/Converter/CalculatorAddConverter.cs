using System;
using System.Globalization;
using System.Windows.Data;

namespace Convert2DTo3D
{
    public class CalculatorAddConverter : IValueConverter
    {
        private int _Step = 0;

        public int Step
        {
            get { return _Step; }
            set { _Step = value; }
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int valueInt = 0;
            if (int.TryParse(value.ToString(), out valueInt))
            {
                valueInt += Step;
                return valueInt;
            }
            ;
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}