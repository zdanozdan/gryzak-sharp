using System;
using System.Globalization;
using System.Windows.Data;

namespace Gryzak.Converters
{
    /// <summary>
    /// Zwraca max(value, ConverterParameter). Używane do Width = max(ViewportWidth, MinWidth tabeli),
    /// żeby przy poziomym scrollu kolumny nie były ściskane poniżej swoich szerokości.
    /// </summary>
    public class MinDoubleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var viewport = value is double d ? d : 0d;
            var min = 0d;
            if (parameter != null)
            {
                double.TryParse(parameter.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out min);
            }

            if (double.IsNaN(viewport) || double.IsInfinity(viewport) || viewport < 0)
            {
                viewport = 0;
            }

            return Math.Max(viewport, min);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
            throw new NotSupportedException();
    }
}
