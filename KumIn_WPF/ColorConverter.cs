using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Globalization;

namespace KumIn_WPF
{
    public class YellowConverter : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, CultureInfo culture)
        {
            int min = 20;
            int max = 30;

            if (int.Parse(value[1].ToString()) == 2)
            {
                min = 40;
                max = 60;
            }

            TimeSpan duration = TimeSpan.ParseExact(value[0].ToString(), "c", CultureInfo.InvariantCulture);
            return duration.TotalMinutes >= min && duration.TotalMinutes < max;
        }

        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Nope.");
        }
    }

    public class RedConverter : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, CultureInfo culture)
        {
            int max = 30;

            if (int.Parse(value[1].ToString()) == 2)
            {
                max = 60;
            }

            TimeSpan duration = TimeSpan.ParseExact(value.ToString(), "c", CultureInfo.InvariantCulture);
            return duration.TotalMinutes >= max;
        }

        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Nope.");
        }
    }
}
