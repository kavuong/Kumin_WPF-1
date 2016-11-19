using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Globalization;

namespace KumIn_WPF
{
    public class YellowConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return TimeSpan.ParseExact(value.ToString(), "h:mm", CultureInfo.InvariantCulture).TotalMinutes > 20
                && TimeSpan.ParseExact(value.ToString(), "h:mm", CultureInfo.InvariantCulture).TotalMinutes < 30;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Nope.");
        }
    }

    public class RedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return TimeSpan.ParseExact(value.ToString(), "h:mm", CultureInfo.InvariantCulture).TotalMinutes >= 30;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Nope.");
        }
    }
}
