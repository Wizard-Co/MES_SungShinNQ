using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WizMes_SungShinNQ
{
    public class DataGridCellFontSizeConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double width)
            {
                double fontSize = width / 8; // 비율 조정
                return Math.Max(8, Math.Min(16, fontSize)); // 8~16 사이로 제한
            }
            return 12; // 기본값
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class DataGridCellMultiFontSizeConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length >= 2 && values[0] is double width && values[1] != null)
            {
                string content = values[1].ToString();

                double baseFontSize = width / 8;
                if (content.Length > 10)
                    baseFontSize *= 0.8;

                return Math.Max(8, Math.Min(16, baseFontSize));
            }
            return 12;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
