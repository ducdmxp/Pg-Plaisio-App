using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pg_Plaisio_App
{
    public static class StringExtension
    {
        public static double ToDouble(this string str)
        {
            if (double.TryParse(str, out double result))
            {
                return result;
            }
            return 0.0;
        }

        public static int ToInt(this string str)
        {
            if (int.TryParse(str, out int result))
            {
                return result;
            }
            return 0;
        }
    }
}