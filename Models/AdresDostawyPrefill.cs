namespace Gryzak.Models
{
    /// <summary>
    /// Dane adresu wysyłki do wstępnego wypełnienia kartoteki kontrahenta w Sferze (przed Wyswietl).
    /// </summary>
    public class AdresDostawyPrefill
    {
        public string? Nazwa { get; set; }
        public string? Ulica { get; set; }
        public string? NrDomu { get; set; }
        public string? NrLokalu { get; set; }
        public string? Kod { get; set; }
        public string? Miejscowosc { get; set; }
        public int? PanstwoId { get; set; }
        public int? WojewodztwoId { get; set; }
    }
}
