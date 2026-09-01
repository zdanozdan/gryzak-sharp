using System.Text.Json.Serialization;

namespace Gryzak.Models
{
    /// <summary>
    /// Body dla PUT /kontrahenci/{id} (częściowe — brak pola = bez zmiany).
    /// </summary>
    public class KontrahentUpdateRequest
    {
        [JsonPropertyName("email")]
        public string? Email { get; set; }

        [JsonPropertyName("adres")]
        public KontrahentAddressUpdate? Adres { get; set; }

        [JsonPropertyName("adresKorespondencyjny")]
        public KontrahentSecondaryAddressUpdate? AdresKorespondencyjny { get; set; }

        [JsonPropertyName("adresDostawy")]
        public KontrahentSecondaryAddressUpdate? AdresDostawy { get; set; }
    }

    public class KontrahentAddressUpdate
    {
        [JsonPropertyName("nazwa")]
        public string? Nazwa { get; set; }

        [JsonPropertyName("nazwaPelna")]
        public string? NazwaPelna { get; set; }

        [JsonPropertyName("nip")]
        public string? Nip { get; set; }

        [JsonPropertyName("ulica")]
        public string? Ulica { get; set; }

        [JsonPropertyName("nrDomu")]
        public string? NrDomu { get; set; }

        [JsonPropertyName("nrLokalu")]
        public string? NrLokalu { get; set; }

        [JsonPropertyName("kod")]
        public string? Kod { get; set; }

        [JsonPropertyName("miejscowosc")]
        public string? Miejscowosc { get; set; }

        [JsonPropertyName("panstwoId")]
        public int? PanstwoId { get; set; }

        [JsonPropertyName("wojewodztwoId")]
        public int? WojewodztwoId { get; set; }
    }

    public class KontrahentSecondaryAddressUpdate
    {
        [JsonPropertyName("aktywny")]
        public bool Aktywny { get; set; }

        [JsonPropertyName("nazwa")]
        public string? Nazwa { get; set; }

        [JsonPropertyName("ulica")]
        public string? Ulica { get; set; }

        [JsonPropertyName("nrDomu")]
        public string? NrDomu { get; set; }

        [JsonPropertyName("nrLokalu")]
        public string? NrLokalu { get; set; }

        [JsonPropertyName("kod")]
        public string? Kod { get; set; }

        [JsonPropertyName("miejscowosc")]
        public string? Miejscowosc { get; set; }

        [JsonPropertyName("panstwoId")]
        public int? PanstwoId { get; set; }

        [JsonPropertyName("wojewodztwoId")]
        public int? WojewodztwoId { get; set; }
    }
}
