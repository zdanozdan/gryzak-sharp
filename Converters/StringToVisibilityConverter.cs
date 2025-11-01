using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Gryzak.Converters
{
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return Visibility.Collapsed;
            
            var str = value.ToString();
            if (string.IsNullOrWhiteSpace(str))
                return Visibility.Collapsed;
            
            // Jeśli wartość to "Brak telefonu" lub "Brak email", ukryj
            if (str == "Brak telefonu" || str == "Brak email")
                return Visibility.Collapsed;
            
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

