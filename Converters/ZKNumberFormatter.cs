using System;
using System.Globalization;
using System.Windows.Data;

namespace Gryzak.Converters
{
    public class ZKNumberFormatter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
            {
                return "";
            }
            
            string number = value.ToString()!.Trim();
            if (string.IsNullOrEmpty(number))
            {
                return "";
            }
            
            return $" ({number})";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

