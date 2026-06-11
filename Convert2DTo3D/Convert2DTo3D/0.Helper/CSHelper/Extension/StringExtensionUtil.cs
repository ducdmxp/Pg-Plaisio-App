using System;
using System.Linq;

namespace Convert2DTo3D
{
    public static class StringExtensionUtil
    {
        public static bool Contains(this string target, string value, StringComparison comparison)
        {
            if (string.IsNullOrEmpty(value) == true)
            {
                return true;
            }
            return target.IndexOf(value, comparison) >= 0;
        }

        public static bool IsUpperCase(this string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                if (Char.IsLetter(input[i]) && !Char.IsUpper(input[i]))
                    return false;
            }
            return true;
        }

        public static bool IsNumberic(this string input)
        {
            return input.All(char.IsDigit);
        }
    }
}