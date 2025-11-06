namespace Gryzak.Models
{
    /// <summary>
    /// Model reprezentujący stawkę VAT z bazy danych Subiekta GT
    /// </summary>
    public class VatRate
    {
        /// <summary>
        /// Identyfikator stawki VAT (vat_id)
        /// </summary>
        public int VatId { get; set; }

        /// <summary>
        /// Symbol stawki VAT (vat_Symbol), np. "23", "8", "5", "0", "ZW", "npo", "PL-23", "CZ-21" itp.
        /// </summary>
        public string VatSymbol { get; set; } = "";

        /// <summary>
        /// Wartość procentowa stawki VAT (vat_Stawka)
        /// </summary>
        public decimal VatStawka { get; set; }
    }
}

