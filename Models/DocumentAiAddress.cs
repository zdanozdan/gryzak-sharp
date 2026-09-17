using System;
using System.Collections.Generic;
using System.Linq;

namespace Gryzak.Models
{
    /// <summary>
    /// Adres wysyłki GLS wyciągnięty z uwag dokumentu przez LLM.
    /// Przechowywany lokalnie w SQLite per (dok_id, dok_typ).
    /// </summary>
    public sealed class DocumentAiAddress
    {
        public int DokId { get; set; }
        public int DokTyp { get; set; }
        public string NrPelny { get; set; } = "";

        public string Name1 { get; set; } = "";
        public string Name2 { get; set; } = "";
        public string Name3 { get; set; } = "";
        public string Street { get; set; } = "";
        public string ZipCode { get; set; } = "";
        public string City { get; set; } = "";
        public string Country { get; set; } = "PL";
        public string Phone { get; set; } = "";
        public string Contact { get; set; } = "";
        public string Notes { get; set; } = "";

        /// <summary>Hash tekstu uwag użytego przy ekstrakcji (wykrycie nieaktualnego wyniku).</summary>
        public string UwagiHash { get; set; } = "";

        /// <summary>Surowy JSON odpowiedzi LLM (debug / ponowne mapowanie).</summary>
        public string RawJson { get; set; } = "";

        public string Confidence { get; set; } = "";
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(Name1)
            && string.IsNullOrWhiteSpace(Street)
            && string.IsNullOrWhiteSpace(ZipCode)
            && string.IsNullOrWhiteSpace(City);

        public bool IsComplete =>
            !string.IsNullOrWhiteSpace(Name1)
            && !string.IsNullOrWhiteSpace(Street)
            && !string.IsNullOrWhiteSpace(ZipCode)
            && !string.IsNullOrWhiteSpace(City);

        public string FormatDisplayName()
        {
            var parts = new[] { Name1, Name2, Name3 }
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Select(p => p.Trim());
            return string.Join(" ", parts);
        }

        public string FormatDisplayAddress()
        {
            var lines = new List<string>();
            if (!string.IsNullOrWhiteSpace(Street))
            {
                lines.Add(Street.Trim());
            }

            var zipCity = $"{ZipCode} {City}".Trim();
            if (!string.IsNullOrWhiteSpace(zipCity))
            {
                lines.Add(zipCity);
            }

            if (!string.IsNullOrWhiteSpace(Country)
                && !Country.Equals("PL", StringComparison.OrdinalIgnoreCase))
            {
                lines.Add(Country.Trim());
            }

            return string.Join(Environment.NewLine, lines);
        }
    }
}
