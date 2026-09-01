using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Gryzak.Models
{
    /// <summary>Lokalny cache GLS (przygotowalnia efemeryczna + nadania read-only).</summary>
    public class GlsShipmentRecord
    {
        public int DokId { get; set; }
        public int DokTyp { get; set; }
        public string NrPelny { get; set; } = "";
        public string Carrier { get; set; } = "GLS";
        public int BoxId { get; set; }
        public string NrPrzyg { get; set; } = "";
        public string NrNad { get; set; } = "";
        public DateTime UpdatedAt { get; set; }

        public string TypKod => SubiektApiDocumentTypes.FromDokTyp(DokTyp);

        public bool HasBoxId => BoxId > 0;

        public bool HasAnyNrListu =>
            !string.IsNullOrWhiteSpace(NrPrzyg) || !string.IsNullOrWhiteSpace(NrNad);

        public bool IsEmpty => !HasBoxId && !HasAnyNrListu;

        public bool IsPreparingBoxOnly =>
            HasBoxId && string.IsNullOrWhiteSpace(NrPrzyg);

        public string SummaryText
        {
            get
            {
                var bits = new List<string>();
                if (HasBoxId)
                {
                    bits.Add($"id {BoxId}");
                }

                if (!string.IsNullOrWhiteSpace(NrPrzyg))
                {
                    bits.Add($"przygotowalnia {NrPrzyg}");
                }

                if (!string.IsNullOrWhiteSpace(NrNad))
                {
                    bits.Add($"nadane {NrNad}");
                }

                return string.Join(", ", bits);
            }
        }

        public string ListDisplayText
        {
            get
            {
                var numbers = CombineNrListu(NrPrzyg, NrNad);
                if (!string.IsNullOrEmpty(numbers))
                {
                    return numbers;
                }

                return HasBoxId ? BoxId.ToString(CultureInfo.InvariantCulture) : "";
            }
        }

        public static bool SameNrListu(string? left, string? right)
        {
            var a = SplitNrListu(left);
            var b = SplitNrListu(right);
            if (a.Count != b.Count)
            {
                return false;
            }

            return a.OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(b.OrderBy(s => s, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase);
        }

        public static List<string> SplitNrListu(string? value)
        {
            var parts = new List<string>();
            if (string.IsNullOrWhiteSpace(value))
            {
                return parts;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in value.Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var part = token.Trim();
                if (part.Length > 0 && seen.Add(part))
                {
                    parts.Add(part);
                }
            }

            return parts;
        }

        public static string NormalizeNrListu(string? value) =>
            string.Join(",", SplitNrListu(value));

        public static string FormatNrListuDisplay(string? value)
        {
            var parts = SplitNrListu(value);
            return parts.Count == 0 ? "" : string.Join(Environment.NewLine, parts);
        }

        public static string CombineNrListu(params string?[] values) =>
            NormalizeNrListu(string.Join(",", values.SelectMany(SplitNrListu)));
    }
}
