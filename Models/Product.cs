using System.Collections.Generic;
using System.Globalization;

namespace Gryzak.Models
{
    public class Product
    {
        public string ProductId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Model { get; set; } = "";
        public int Quantity { get; set; }
        public string Price { get; set; } = "";
        public string Total { get; set; } = "";
        public string Tax { get; set; } = "";
        public string TaxRate { get; set; } = "";
        public string Discount { get; set; } = "";
        public List<ProductOption>? Options { get; set; }
        
        // Wartość brutto dla pozycji = (Price + Tax) * Quantity
        public string GrossTotal
        {
            get
            {
                var price = 0.0;
                var tax = 0.0;
                
                if (!string.IsNullOrWhiteSpace(Price))
                {
                    double.TryParse(Price, NumberStyles.Float, CultureInfo.InvariantCulture, out price);
                }
                
                if (!string.IsNullOrWhiteSpace(Tax))
                {
                    double.TryParse(Tax, NumberStyles.Float, CultureInfo.InvariantCulture, out tax);
                }
                
                var gross = (price + tax) * Quantity;
                return gross.ToString("F2", CultureInfo.InvariantCulture);
            }
        }
    }

    public class ProductOption
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public class OrderTotal
    {
        public string Code { get; set; } = "";
        public string Title { get; set; } = "";
        public double Value { get; set; } = 0.0;
        public int SortOrder { get; set; } = 0;
        
        public string FormattedValue => Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
        
        public bool IsTotal => Code == "total";
        public bool IsDiscount => Code == "coupon" || Value < 0;
    }
}

