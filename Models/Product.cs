using System.Collections.Generic;

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
    }

    public class ProductOption
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
    }
}

