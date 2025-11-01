using System;
using System.Globalization;
using System.Windows.Data;

namespace Gryzak.Converters
{
    public class GrossPriceConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2)
                return "-";

            if (values[0] == null)
                return "-";

            var priceStr = values[0].ToString();
            if (string.IsNullOrWhiteSpace(priceStr))
                return "-";

            if (!double.TryParse(priceStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var price))
                return "-";

            var tax = 0.0;
            if (values[1] != null)
            {
                var taxStr = values[1].ToString();
                if (!string.IsNullOrWhiteSpace(taxStr))
                {
                    double.TryParse(taxStr, NumberStyles.Float, CultureInfo.InvariantCulture, out tax);
                }
            }

            var currency = "PLN";
            if (values.Length > 2 && values[2] != null)
            {
                var currencyStr = values[2].ToString();
                if (!string.IsNullOrWhiteSpace(currencyStr))
                {
                    currency = currencyStr;
                }
            }

            var grossPrice = price + tax;
            return grossPrice.ToString("F2", CultureInfo.InvariantCulture) + " " + currency;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

