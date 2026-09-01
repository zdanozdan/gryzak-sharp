using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Gryzak.Models
{
    public class GlsParcel
    {
        public string Weight { get; set; } = "1";
        public string Reference { get; set; } = "";
        public string Number { get; set; } = "";
    }

    public class GlsConsignment
    {
        public const string DefaultParcelWeightKg = "1";
        public const int NameFieldMaxLength = 40;
        public const int MaxParcelCount = 99;
        public const int GlsReferenceMaxLength = 25;

        private static readonly Regex ReferenceDokIdRegex = new(
            @"\(#\s*(\d+)\s*\)",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public string Name1 { get; set; } = "";
        public string Name2 { get; set; } = "";
        public string Name3 { get; set; } = "";
        public string Country { get; set; } = "PL";
        public string ZipCode { get; set; } = "";
        public string City { get; set; } = "";
        public string Street { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Contact { get; set; } = "";
        public string References { get; set; } = "";
        public string Notes { get; set; } = "";
        public bool CashOnDelivery { get; set; }
        public decimal CodAmount { get; set; }
        public int? ExistingId { get; set; }
        public List<GlsParcel> Parcels { get; set; } = new();

        public bool IsEdit => ExistingId is > 0;
        public string CodAmountDisplay => CodAmount.ToString("N2", CultureInfo.GetCultureInfo("pl-PL"));

        public string SummaryAddress =>
            string.Join(", ", new[] { Street, $"{ZipCode} {City}".Trim(), Country }
                .Where(s => !string.IsNullOrWhiteSpace(s)));

        public string ParcelWeight
        {
            get => Parcels.FirstOrDefault()?.Weight ?? DefaultParcelWeightKg;
            set
            {
                EnsureParcel();
                foreach (var parcel in Parcels)
                {
                    parcel.Weight = value;
                }
            }
        }

        public int ParcelCount => Math.Max(1, Parcels.Count);

        public void SetParcelCount(int count)
        {
            if (count < 1)
            {
                count = 1;
            }

            if (count > MaxParcelCount)
            {
                count = MaxParcelCount;
            }

            EnsureParcel();
            var template = Parcels[0];
            var weight = string.IsNullOrWhiteSpace(template.Weight) ? DefaultParcelWeightKg : template.Weight;
            var reference = string.IsNullOrWhiteSpace(template.Reference) ? References : template.Reference;

            while (Parcels.Count < count)
            {
                Parcels.Add(new GlsParcel
                {
                    Weight = weight,
                    Reference = reference
                });
            }

            while (Parcels.Count > count)
            {
                Parcels.RemoveAt(Parcels.Count - 1);
            }

            foreach (var parcel in Parcels)
            {
                if (string.IsNullOrWhiteSpace(parcel.Weight))
                {
                    parcel.Weight = weight;
                }

                if (string.IsNullOrWhiteSpace(parcel.Reference))
                {
                    parcel.Reference = reference;
                }
            }
        }

        public static GlsConsignment FromDocument(SubiektDocument? document)
        {
            var consignment = new GlsConsignment();
            if (document == null)
            {
                consignment.Parcels.Add(new GlsParcel { Weight = DefaultParcelWeightKg });
                return consignment;
            }

            var useDelivery = document.HasAdresDostawy;
            var name = (document.KontrahentNazwa ?? "").Trim();
            SplitName(name, out var name1, out var name2, out var name3);

            var street = useDelivery
                ? Truncate((document.AdresDostawyUlica ?? "").Trim(), 40)
                : Truncate((document.KontrahentUlica ?? "").Trim(), 40);
            var zip = useDelivery
                ? NormalizeZip(document.AdresDostawyKodPocztowy)
                : NormalizeZip(document.KontrahentKodPocztowy);
            var city = useDelivery
                ? Truncate((document.AdresDostawyMiejscowosc ?? "").Trim(), 40)
                : Truncate((document.KontrahentMiejscowosc ?? "").Trim(), 40);

            if (string.IsNullOrWhiteSpace(street) || string.IsNullOrWhiteSpace(zip) || string.IsNullOrWhiteSpace(city))
            {
                var fallbackAddress = useDelivery ? document.AdresDostawyAdres : document.KontrahentAdres;
                TryParseAddress(fallbackAddress, ref street, ref zip, ref city);
            }

            if (useDelivery
                && (string.IsNullOrWhiteSpace(street) || string.IsNullOrWhiteSpace(zip) || string.IsNullOrWhiteSpace(city)))
            {
                // Niepełny adres wysyłki — dopełnij z adresu podstawowego kontrahenta.
                if (string.IsNullOrWhiteSpace(street))
                {
                    street = Truncate((document.KontrahentUlica ?? "").Trim(), 40);
                }

                if (string.IsNullOrWhiteSpace(zip))
                {
                    zip = NormalizeZip(document.KontrahentKodPocztowy);
                }

                if (string.IsNullOrWhiteSpace(city))
                {
                    city = Truncate((document.KontrahentMiejscowosc ?? "").Trim(), 40);
                }

                if (string.IsNullOrWhiteSpace(street) || string.IsNullOrWhiteSpace(zip) || string.IsNullOrWhiteSpace(city))
                {
                    TryParseAddress(document.KontrahentAdres, ref street, ref zip, ref city);
                }
            }

            var references = BuildGlsReference(document);

            var phone = useDelivery && !string.IsNullOrWhiteSpace(document.AdresDostawyTelefon)
                ? document.AdresDostawyTelefon.Trim()
                : (document.KontrahentTelefon ?? "").Trim();
            var country = useDelivery && !string.IsNullOrWhiteSpace(document.AdresDostawyKrajKod)
                ? document.AdresDostawyKrajKod
                : document.KontrahentKrajKod;

            consignment.Name1 = name1;
            consignment.Name2 = name2;
            consignment.Name3 = name3;
            consignment.Country = NormalizeCountry(country);
            consignment.ZipCode = zip;
            consignment.City = city;
            consignment.Street = street;
            consignment.Phone = Truncate(phone, 25);
            consignment.Contact = Truncate((document.KontrahentEmail ?? "").Trim(), 40);
            consignment.References = references;
            consignment.Notes = Truncate(
                string.IsNullOrWhiteSpace(document.NrPelnyOryg)
                    ? ""
                    : document.NrPelnyOryg.Trim(),
                40);
            consignment.CodAmount = document.WartBrutto > 0 ? document.WartBrutto : 0;
            consignment.CashOnDelivery = document.IsGlsPobranie;
            consignment.Parcels.Add(new GlsParcel
            {
                Weight = DefaultParcelWeightKg,
                Reference = references
            });

            return consignment;
        }

        public static bool TryFromDocument(SubiektDocument document, out GlsConsignment consignment, out string error)
        {
            consignment = FromDocument(document);
            return consignment.TryValidate(out error);
        }

        public bool TryValidate(out string error)
        {
            error = "";
            NormalizeLimits();

            if (string.IsNullOrWhiteSpace(Name1))
            {
                error = "Brak nazwy odbiorcy — GLS wymaga pola rname1.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(Street))
            {
                error = "Brak ulicy odbiorcy — GLS wymaga pola rstreet.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(ZipCode))
            {
                error = "Brak kodu pocztowego odbiorcy — GLS wymaga pola rzipcode.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(City))
            {
                error = "Brak miejscowości odbiorcy — GLS wymaga pola rcity.";
                return false;
            }

            if (CashOnDelivery && CodAmount <= 0)
            {
                error = "Za pobraniem wymaga kwoty większej od zera (wartość brutto dokumentu).";
                return false;
            }

            EnsureParcel();
            if (Parcels.Count < 1 || Parcels.Count > MaxParcelCount)
            {
                error = $"Ilość paczek musi być w zakresie 1–{MaxParcelCount}.";
                return false;
            }

            return true;
        }

        public GlsConsignment Clone()
        {
            return new GlsConsignment
            {
                Name1 = Name1,
                Name2 = Name2,
                Name3 = Name3,
                Country = Country,
                ZipCode = ZipCode,
                City = City,
                Street = Street,
                Phone = Phone,
                Contact = Contact,
                References = References,
                Notes = Notes,
                CashOnDelivery = CashOnDelivery,
                CodAmount = CodAmount,
                ExistingId = ExistingId,
                Parcels = Parcels.Select(p => new GlsParcel
                {
                    Weight = p.Weight,
                    Reference = p.Reference,
                    Number = p.Number
                }).ToList()
            };
        }

        public string GetParcelNumbersDisplay()
        {
            var numbers = Parcels
                .Select(p => (p.Number ?? "").Trim())
                .Where(n => !string.IsNullOrEmpty(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            return string.Join(", ", numbers);
        }

        public void NormalizeLimits()
        {
            Name1 = Truncate(Name1, NameFieldMaxLength);
            Name2 = Truncate(Name2, NameFieldMaxLength);
            Name3 = Truncate(Name3, NameFieldMaxLength);
            Country = NormalizeCountry(Country);
            ZipCode = NormalizeZip(ZipCode);
            City = Truncate(City, 40);
            Street = Truncate(Street, 40);
            Phone = Truncate(Phone, 25);
            Contact = Truncate(Contact, 40);
            References = Truncate(References, GlsReferenceMaxLength);
            Notes = Truncate(Notes, 40);

            EnsureParcel();
            if (Parcels.Count > MaxParcelCount)
            {
                Parcels.RemoveRange(MaxParcelCount, Parcels.Count - MaxParcelCount);
            }

            foreach (var parcel in Parcels)
            {
                parcel.Reference = Truncate(parcel.Reference, GlsReferenceMaxLength);
                if (string.IsNullOrWhiteSpace(parcel.Weight))
                {
                    parcel.Weight = DefaultParcelWeightKg;
                }
            }
        }

        private void EnsureParcel()
        {
            if (Parcels.Count == 0)
            {
                Parcels.Add(new GlsParcel
                {
                    Weight = DefaultParcelWeightKg,
                    Reference = References
                });
            }
        }

        private static string NormalizeCountry(string? value)
        {
            var country = (value ?? "").Trim().ToUpperInvariant();
            if (country.Length == 2 && country.All(char.IsLetter))
            {
                return country;
            }

            if (country is "POLSKA" or "POLAND" or "PLN")
            {
                return "PL";
            }

            return "PL";
        }

        private static string NormalizeZip(string? value)
        {
            var zip = (value ?? "").Trim();
            if (string.IsNullOrWhiteSpace(zip))
            {
                return "";
            }

            var digits = new string(zip.Where(char.IsDigit).ToArray());
            if (digits.Length == 5)
            {
                return $"{digits[..2]}-{digits[2..]}";
            }

            return Truncate(zip.Replace(" ", ""), 10);
        }

        private static void TryParseAddress(string? address, ref string street, ref string zip, ref string city)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return;
            }

            var text = address.Trim();
            var match = Regex.Match(text, @"^(.*?)[,\s]+(\d{2}[-\s]?\d{3})\s+(.+)$");
            if (!match.Success)
            {
                match = Regex.Match(text, @"(\d{2}[-\s]?\d{3})\s+(.+)$");
                if (match.Success)
                {
                    if (string.IsNullOrWhiteSpace(zip))
                    {
                        zip = NormalizeZip(match.Groups[1].Value);
                    }

                    if (string.IsNullOrWhiteSpace(city))
                    {
                        city = Truncate(match.Groups[2].Value.Trim().Trim(','), 40);
                    }

                    if (string.IsNullOrWhiteSpace(street))
                    {
                        street = Truncate(text[..match.Index].Trim().Trim(','), 40);
                    }
                }

                return;
            }

            if (string.IsNullOrWhiteSpace(street))
            {
                street = Truncate(match.Groups[1].Value.Trim().Trim(','), 40);
            }

            if (string.IsNullOrWhiteSpace(zip))
            {
                zip = NormalizeZip(match.Groups[2].Value);
            }

            if (string.IsNullOrWhiteSpace(city))
            {
                city = Truncate(match.Groups[3].Value.Trim().Trim(','), 40);
            }
        }

        public static void SplitName(string? name, out string name1, out string name2, out string name3)
        {
            name1 = "";
            name2 = "";
            name3 = "";

            var remaining = Regex.Replace((name ?? "").Trim(), @"\s+", " ");
            if (string.IsNullOrEmpty(remaining))
            {
                return;
            }

            name1 = TakeNamePart(ref remaining);
            name2 = TakeNamePart(ref remaining);
            name3 = TakeNamePart(ref remaining);
        }

        private static string TakeNamePart(ref string remaining)
        {
            if (string.IsNullOrEmpty(remaining))
            {
                return "";
            }

            if (remaining.Length <= NameFieldMaxLength)
            {
                var part = remaining;
                remaining = "";
                return part;
            }

            var window = remaining[..NameFieldMaxLength];
            var lastSpace = window.LastIndexOf(' ');
            if (lastSpace > 0)
            {
                var part = remaining[..lastSpace].TrimEnd();
                remaining = remaining[(lastSpace + 1)..].TrimStart();
                return part;
            }

            var hardCut = remaining[..NameFieldMaxLength];
            remaining = remaining[NameFieldMaxLength..].TrimStart();
            return hardCut;
        }

        private static string Truncate(string? value, int maxLength)
        {
            var text = (value ?? "").Trim();
            return text.Length <= maxLength ? text : text[..maxLength].Trim();
        }

        /// <summary>
        /// Referencja GLS: skrócony numer dokumentu + stabilny sufiks <c>(#dok_id)</c>.
        /// </summary>
        public static string BuildGlsReference(SubiektDocument? document)
        {
            if (document == null)
            {
                return "";
            }

            var nr = !string.IsNullOrWhiteSpace(document.NrPelny)
                ? document.NrPelny.Trim()
                : (document.NrPelnyOryg ?? "").Trim();
            if (document.DokId <= 0)
            {
                return "";
            }

            var suffix = $" (#{document.DokId})";
            var maxNrLen = GlsReferenceMaxLength - suffix.Length;
            if (maxNrLen < 1)
            {
                return Truncate(suffix, GlsReferenceMaxLength);
            }

            return Truncate(nr, maxNrLen) + suffix;
        }

        public static bool TryParseReferenceDokId(string? references, out int dokId)
        {
            dokId = 0;
            var match = ReferenceDokIdRegex.Match(references ?? "");
            if (!match.Success)
            {
                return false;
            }

            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out dokId)
                   && dokId > 0;
        }

        public static bool MatchesDocumentReference(string? references, SubiektDocument? document)
        {
            if (document == null || !TryParseReferenceDokId(references, out var refDokId))
            {
                return false;
            }

            if (document.DokId > 0 && refDokId == document.DokId)
            {
                return true;
            }

            return document.GlsShipmentDokId is int owner && owner == refDokId;
        }

        public static SubiektDocument? FindDocumentForReferenceDokId(
            string? references,
            IEnumerable<SubiektDocument> documents)
        {
            if (!TryParseReferenceDokId(references, out var refDokId))
            {
                return null;
            }

            foreach (var doc in documents)
            {
                if (doc.DokId == refDokId)
                {
                    return doc;
                }

                if (doc.GlsShipmentDokId is int owner && owner == refDokId)
                {
                    return doc;
                }
            }

            return null;
        }

        /// <summary>Dopasowanie po fragmencie numeru dokumentu w referencji GLS (gdy brak trafienia po dok_id).</summary>
        public static SubiektDocument? FindDocumentForReferenceNrPelny(
            string? references,
            IEnumerable<SubiektDocument> documents)
        {
            var refs = (references ?? "").Trim();
            if (refs.Length == 0)
            {
                return null;
            }

            SubiektDocument? best = null;
            var bestLen = 0;
            foreach (var doc in documents)
            {
                var nr = (doc.NrPelny ?? "").Trim();
                if (nr.Length == 0 || bestLen >= nr.Length)
                {
                    continue;
                }

                if (refs.Contains(nr, StringComparison.OrdinalIgnoreCase))
                {
                    best = doc;
                    bestLen = nr.Length;
                }
            }

            return best;
        }
    }

    public class GlsCreateShipmentResult
    {
        public bool Success { get; set; }
        public int? ConsignmentId { get; set; }
        public string? ErrorMessage { get; set; }
        public string? WarningMessage { get; set; }
        public bool NotFound { get; set; }
        public string EnvironmentName { get; set; } = "";
    }

    public class GlsGetShipmentResult
    {
        public bool Success { get; set; }
        public bool NotFound { get; set; }
        public GlsConsignment? Consignment { get; set; }
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
    }

    public class GlsPreparingBoxItem
    {
        public int Id { get; set; }
        public string References { get; set; } = "";
        public string ParcelNumber { get; set; } = "";

        public string Title
        {
            get
            {
                var numbers = SubiektPrzesylka.FormatNrListuDisplay(ParcelNumber);
                return string.IsNullOrWhiteSpace(numbers)
                    ? Id.ToString(CultureInfo.InvariantCulture)
                    : numbers;
            }
        }

        public string Subtitle
        {
            get
            {
                var idPart = $"id {Id.ToString(CultureInfo.InvariantCulture)}";
                if (!string.IsNullOrWhiteSpace(ParcelNumber))
                {
                    return string.IsNullOrWhiteSpace(References) ? idPart : $"{idPart}  •  {References}";
                }

                return References;
            }
        }
    }

    public class GlsPickupItem
    {
        public int Id { get; set; }
        public int PickupId { get; set; }
        public string References { get; set; } = "";
        public string ParcelNumber { get; set; } = "";

        public string Title
        {
            get
            {
                var numbers = SubiektPrzesylka.FormatNrListuDisplay(ParcelNumber);
                return string.IsNullOrWhiteSpace(numbers)
                    ? Id.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : numbers;
            }
        }

        public string Subtitle
        {
            get
            {
                var idPart = $"id {Id.ToString(System.Globalization.CultureInfo.InvariantCulture)}";
                return string.IsNullOrWhiteSpace(References) ? idPart : $"{References}  •  {idPart}";
            }
        }
    }

    public class GlsStatusListsResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
        public List<GlsPreparingBoxItem> PreparingBoxItems { get; set; } = new();
        public List<GlsPickupItem> PickupItems { get; set; } = new();
        public string? PickupErrorMessage { get; set; }
    }

    public class GlsPreparingBoxListResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
        public List<GlsPreparingBoxItem> Items { get; set; } = new();
    }

    public enum GlsPickupRangeMode
    {
        Today = 0,
        TwoDays = 1,
        ThreeDays = 2,
        Week = 3,
        ThirtyDays = 4,
        CustomDay = 5,
        CustomRange = 6
    }

    public class GlsPickupQuery
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }

        public static GlsPickupQuery ForLookbackDays(int lookbackDaysInclusive)
        {
            var today = DateTime.Now.Date;
            var days = Math.Max(1, lookbackDaysInclusive);
            return new GlsPickupQuery
            {
                FromDate = today.AddDays(-(days - 1)),
                ToDate = today
            };
        }

        public static GlsPickupQuery ForSingleDay(DateTime day) =>
            new()
            {
                FromDate = day.Date,
                ToDate = day.Date
            };

        public static GlsPickupQuery ForDateRange(DateTime from, DateTime to) =>
            new()
            {
                FromDate = from.Date,
                ToDate = to.Date
            };

        public static GlsPickupQuery FromRangeMode(
            GlsPickupRangeMode mode,
            DateTime? customDay = null,
            DateTime? customFrom = null,
            DateTime? customTo = null)
        {
            return mode switch
            {
                GlsPickupRangeMode.TwoDays => ForLookbackDays(2),
                GlsPickupRangeMode.ThreeDays => ForLookbackDays(3),
                GlsPickupRangeMode.Week => ForLookbackDays(7),
                GlsPickupRangeMode.ThirtyDays => ForLookbackDays(30),
                GlsPickupRangeMode.CustomDay => ForSingleDay(customDay ?? DateTime.Now.Date),
                GlsPickupRangeMode.CustomRange => ForDateRange(
                    customFrom ?? DateTime.Now.Date,
                    customTo ?? DateTime.Now.Date),
                _ => ForLookbackDays(1)
            };
        }

        public void Normalize()
        {
            FromDate = FromDate.Date;
            ToDate = ToDate.Date;
            if (ToDate < FromDate)
            {
                (FromDate, ToDate) = (ToDate, FromDate);
            }
        }
    }

    public class GlsPickupListResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
        public List<GlsPickupItem> Items { get; set; } = new();
        public bool HitPickupScanLimit { get; set; }
    }

    public class GlsFetchProgress
    {
        public string Operation { get; set; } = "GLS";
        public string Phase { get; set; } = "";
        public int Done { get; set; }
        public int Total { get; set; }
        public int? CurrentId { get; set; }
        public string? CurrentLabel { get; set; }

        public double Percent =>
            Total > 0
                ? Math.Min(100, Math.Round(100.0 * Done / Total, 1))
                : 0;

        public string StatusText
        {
            get
            {
                var parts = new List<string> { Operation };
                if (!string.IsNullOrWhiteSpace(Phase))
                {
                    parts.Add(Phase);
                }

                if (Total > 0)
                {
                    parts.Add($"{Percent:0.#}% ({Done}/{Total})");
                }
                else if (Done > 0)
                {
                    parts.Add($"przetworzono {Done}");
                }

                if (CurrentId is int id && id > 0)
                {
                    parts.Add($"id={id}");
                }

                if (!string.IsNullOrWhiteSpace(CurrentLabel))
                {
                    parts.Add(CurrentLabel!);
                }

                return string.Join(" · ", parts);
            }
        }
    }

    public class GlsLabelResult
    {
        public const string PdfMode = "one_label_on_a4_pdf";
        /// <summary>PDF etykiety 160×100 mm — lepszy do podglądu niż A4.</summary>
        public const string PreviewPdfMode = "roll_160x100_pdf";
        public const string ZebraZplMode = "roll_160x100_zebra";
        public const string DefaultMode = PdfMode;

        public bool Success { get; set; }
        public bool NotFound { get; set; }
        public byte[]? LabelBytes { get; set; }
        public string Mode { get; set; } = DefaultMode;
        public string FileExtension { get; set; } = ".pdf";
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
    }
}
