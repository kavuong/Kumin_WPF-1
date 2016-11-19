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
            TimeSpan duration = TimeSpan.ParseExact(value.ToString(), "c", CultureInfo.InvariantCulture);
            return duration.TotalMinutes >= 1 && duration.TotalMinutes < 2;
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
            TimeSpan duration = TimeSpan.ParseExact(value.ToString(), "c", CultureInfo.InvariantCulture);
            return duration.TotalMinutes >= 2;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("Nope.");
        }
    }
}
