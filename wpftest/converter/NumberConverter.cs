using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WizMes_SungShinNQ
{
    public class NumberConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // parameter로 소숫점 자릿수 지정 (기본값: 0)
            int decimalPlaces = 0;
            if (parameter != null && int.TryParse(parameter.ToString(), out int places))
            {
                decimalPlaces = places;
            }

            string format = decimalPlaces == 0 ? "N0" : $"N{decimalPlaces}";

            if (value is decimal decimalValue)
                return decimalValue.ToString(format);
            if (value is double doubleValue)
                return doubleValue.ToString(format);
            if (value is float floatValue)
                return floatValue.ToString(format);
            if (value is int intValue)
                return intValue.ToString(format);
            if (value is long longValue)
                return longValue.ToString(format);

            return value?.ToString() ?? "";
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue)
            {
                string cleanValue = stringValue.Replace(",", "").Trim();

                int decimalPlaces = 0;
                if (parameter != null && int.TryParse(parameter.ToString(), out int places))
                {
                    decimalPlaces = places;
                }

                if (decimalPlaces == 0)
                {
                    cleanValue = cleanValue.Replace(".", "");
                }

                if (double.TryParse(cleanValue, out double result))
                {
                    // targetType에 따라 적절한 타입으로 반환
                    if (targetType == typeof(double) || targetType == typeof(double?))
                        return result;
                    if (targetType == typeof(decimal) || targetType == typeof(decimal?))
                        return (decimal)result;
                    if (targetType == typeof(float) || targetType == typeof(float?))
                        return (float)result;
                    if (targetType == typeof(int) || targetType == typeof(int?))
                        return (int)result;

                    return result;
                }
            }
            return 0.0;
        }
    }
}