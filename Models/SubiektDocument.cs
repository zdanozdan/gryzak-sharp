using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

namespace Gryzak.Models
{
    public class SubiektDocument : INotifyPropertyChanged
    {
        private bool _isSelected;

        public int DokId { get; set; }
        public int DokTyp { get; set; }
        public string TypKod { get; set; } = "zk";
        public string NrPelny { get; set; } = "";
        public string NrPelnyOryg { get; set; } = "";
        public int? DoDokId { get; set; }
        public string DoDokNrPelny { get; set; } = "";
        public DateTime? DoDokDataWyst { get; set; }
        public bool HasDoDokNrPelny => SubiektDocumentNumber.IsValid(DoDokNrPelny);
        public DateTime? DataWyst { get; set; }
        public decimal WartNetto { get; set; }
        public decimal WartBrutto { get; set; }
        public string StatusNazwa { get; set; } = "";
        public string Wystawil { get; set; } = "";
        public string PlatNazwa { get; set; } = "";
        public string KartaNazwa { get; set; } = "";

        private SubiektPrzesylka? _przesylka;
        private GlsShipmentRecord? _glsShipment;
        private int? _glsPreparingBoxId;
        private string _glsPreparingBoxParcelNumber = "";
        private bool _glsStatusChecked;
        private int? _glsPickupConsignmentId;
        private string _glsPickupParcelNumber = "";

        public SubiektPrzesylka? Przesylka
        {
            get => _przesylka;
            set
            {
                _przesylka = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PrzesylkaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
                OnPropertyChanged(nameof(NadaniaDisplay));
                OnPropertyChanged(nameof(NadaniaTooltip));
            }
        }

        /// <summary>
        /// Dokument-właściciel pola <c>Przesylka</c> (WZ lub FS). Może różnić się od
        /// <see cref="DokId"/>, gdy UI pokazuje przesyłkę z powiązanego WZ na FS/ZK.
        /// </summary>
        public int? PrzesylkaOwnerDokId { get; set; }

        /// <summary>Typ właściciela: <c>wz</c> / <c>fs</c>.</summary>
        public string PrzesylkaOwnerTypKod { get; set; } = "";

        public bool OwnsPrzesylkaField =>
            PrzesylkaOwnerDokId is int owner
                ? owner == DokId
                : SubiektApiDocumentTypes.IsFsOrWz(DokTyp)
                  || SubiektApiDocumentTypes.IsFsOrWz(TypKod);

        /// <summary>Lokalny cache GLS (SQLite), nie pole Subiekta.</summary>
        public GlsShipmentRecord? GlsShipment
        {
            get => _glsShipment;
            set
            {
                _glsShipment = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(GlsShipmentDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
                OnPropertyChanged(nameof(NadaniaDisplay));
                OnPropertyChanged(nameof(NadaniaTooltip));
                OnPropertyChanged(nameof(HasNadaneNrListu));
                OnPropertyChanged(nameof(IsInGlsPreparingBox));
            }
        }

        /// <summary>
        /// Dokument, spod którego wczytano cache (może być powiązane WZ przy FS/ZK).
        /// </summary>
        public int? GlsShipmentDokId { get; set; }

        public int? GlsShipmentDokTyp { get; set; }

        public bool OwnsGlsShipmentCache =>
            GlsShipmentDokId is int owner
                ? owner == DokId
                : true;

        public int? GlsPreparingBoxId
        {
            get => _glsPreparingBoxId;
            set
            {
                if (_glsPreparingBoxId == value) return;
                _glsPreparingBoxId = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
                OnPropertyChanged(nameof(IsInGlsPreparingBox));
            }
        }

        public string GlsPreparingBoxParcelNumber
        {
            get => _glsPreparingBoxParcelNumber;
            set
            {
                var next = value ?? "";
                if (_glsPreparingBoxParcelNumber == next) return;
                _glsPreparingBoxParcelNumber = next;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
            }
        }

        public bool IsInGlsPreparingBox => GlsPreparingBoxId is > 0;

        public bool GlsStatusChecked
        {
            get => _glsStatusChecked;
            set
            {
                if (_glsStatusChecked == value) return;
                _glsStatusChecked = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
            }
        }

        public string PrzygotowalniaDisplay
        {
            get
            {
                // Live GLS (zielony w UI) ma pierwszeństwo; inaczej lokalny cache (szary).
                if (GlsPreparingBoxId is int glsId && glsId > 0)
                {
                    if (!string.IsNullOrWhiteSpace(GlsPreparingBoxParcelNumber))
                    {
                        return GlsShipmentRecord.FormatNrListuDisplay(GlsPreparingBoxParcelNumber);
                    }

                    return glsId.ToString(CultureInfo.InvariantCulture);
                }

                var cached = GlsShipment;
                var cachedHasBox = cached is { HasBoxId: true };

                // Gdy przesyłka opuściła przygotowalnię (brak box_id), numery są w kolumnie Nadania.
                if (!cachedHasBox && HasNadaneNrListu)
                {
                    return "";
                }

                if (cached == null || cached.IsEmpty)
                {
                    return "";
                }

                if (!string.IsNullOrWhiteSpace(cached.NrPrzyg))
                {
                    return GlsShipmentRecord.FormatNrListuDisplay(cached.NrPrzyg);
                }

                return cached.HasBoxId
                    ? cached.BoxId.ToString(CultureInfo.InvariantCulture)
                    : "";
            }
        }

        public string PrzygotowalniaTooltip
        {
            get
            {
                if (GlsPreparingBoxId is int glsId && glsId > 0)
                {
                    var numbers = GlsShipmentRecord.SplitNrListu(GlsPreparingBoxParcelNumber);
                    var idPart = $"id {glsId.ToString(CultureInfo.InvariantCulture)}";
                    return numbers.Count == 0
                        ? idPart
                        : idPart + "\n" + string.Join("\n", numbers);
                }

                var cached = GlsShipment;
                var cachedHasBox = cached is { HasBoxId: true };
                if (!cachedHasBox && HasNadaneNrListu)
                {
                    return "";
                }

                if (cached == null || cached.IsEmpty)
                {
                    return "";
                }

                var cachedNumbers = GlsShipmentRecord.SplitNrListu(cached.NrPrzyg);
                if (cached.HasBoxId)
                {
                    var idPart = $"id {cached.BoxId.ToString(CultureInfo.InvariantCulture)} (cache)";
                    return cachedNumbers.Count == 0
                        ? idPart
                        : idPart + "\n" + string.Join("\n", cachedNumbers);
                }

                return PrzygotowalniaDisplay;
            }
        }

        public int? GlsPickupConsignmentId
        {
            get => _glsPickupConsignmentId;
            set
            {
                if (_glsPickupConsignmentId == value) return;
                _glsPickupConsignmentId = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NadaniaDisplay));
                OnPropertyChanged(nameof(NadaniaTooltip));
                OnPropertyChanged(nameof(IsGlsPickedUp));
            }
        }

        public string GlsPickupParcelNumber
        {
            get => _glsPickupParcelNumber;
            set
            {
                var next = value ?? "";
                if (_glsPickupParcelNumber == next) return;
                _glsPickupParcelNumber = next;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NadaniaDisplay));
                OnPropertyChanged(nameof(NadaniaTooltip));
                OnPropertyChanged(nameof(IsGlsPickedUp));
                OnPropertyChanged(nameof(HasNadaneNrListu));
                OnPropertyChanged(nameof(PrzygotowalniaDisplay));
                OnPropertyChanged(nameof(PrzygotowalniaTooltip));
                OnPropertyChanged(nameof(IsInGlsPreparingBox));
            }
        }

        /// <summary>Przesyłka ma już numer po nadaniu (live lub cache SQLite).</summary>
        public bool HasNadaneNrListu =>
            !string.IsNullOrWhiteSpace(GlsPickupParcelNumber)
            || GlsShipment is { IsEmpty: false } shipment && !string.IsNullOrWhiteSpace(shipment.NrNad);

        /// <summary>
        /// True gdy nadanie potwierdzone z API GLS (zielony w UI).
        /// Same dane z lokalnego cache → szary (<see cref="NadaniaDisplay"/>).
        /// </summary>
        public bool IsGlsPickedUp =>
            GlsPickupConsignmentId is > 0 || !string.IsNullOrWhiteSpace(GlsPickupParcelNumber);

        public string NadaniaDisplay
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(GlsPickupParcelNumber))
                {
                    return GlsShipmentRecord.FormatNrListuDisplay(GlsPickupParcelNumber);
                }

                var nrListu = (GlsShipment?.NrNad ?? "").Trim();
                if (!string.IsNullOrEmpty(nrListu))
                {
                    return GlsShipmentRecord.FormatNrListuDisplay(nrListu);
                }

                return GlsPickupConsignmentId is int id && id > 0
                    ? id.ToString(CultureInfo.InvariantCulture)
                    : "";
            }
        }

        public string NadaniaTooltip
        {
            get
            {
                var parts = GlsShipmentRecord.SplitNrListu(
                    !string.IsNullOrWhiteSpace(GlsPickupParcelNumber)
                        ? GlsPickupParcelNumber
                        : GlsShipment?.NrNad);
                if (parts.Count == 0)
                {
                    return NadaniaDisplay;
                }

                return parts.Count == 1
                    ? parts[0]
                    : $"{parts.Count}× GLS:\n" + string.Join("\n", parts);
            }
        }

        public int KontrahentId { get; set; }
        public string KontrahentSymbol { get; set; } = "";
        public string KontrahentNazwa { get; set; } = "";
        public string KontrahentNip { get; set; } = "";
        public string KontrahentEmail { get; set; } = "";
        public string KontrahentAdres { get; set; } = "";
        public string KontrahentUlica { get; set; } = "";
        public string KontrahentKodPocztowy { get; set; } = "";
        public string KontrahentMiejscowosc { get; set; } = "";
        public string KontrahentKrajKod { get; set; } = "";
        public string KontrahentTelefon { get; set; } = "";

        public bool MatchesSearch(string? query)
        {
            var q = (query ?? "").Trim();
            if (string.IsNullOrEmpty(q))
            {
                return true;
            }

            if (DokId.ToString(CultureInfo.InvariantCulture).Contains(q, StringComparison.OrdinalIgnoreCase)
                || KontrahentId.ToString(CultureInfo.InvariantCulture).Contains(q, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (ContainsIgnoreCase(NrPelny, q)
                || ContainsIgnoreCase(NrPelnyOryg, q)
                || ContainsIgnoreCase(DoDokNrPelny, q)
                || ContainsIgnoreCase(KontrahentSymbol, q)
                || ContainsIgnoreCase(KontrahentNazwa, q)
                || ContainsIgnoreCase(KontrahentEmail, q)
                || ContainsIgnoreCase(KontrahentMiejscowosc, q)
                || ContainsIgnoreCase(KontrahentTelefon, q)
                || ContainsIgnoreCase(PlatnikNazwa, q)
                || ContainsIgnoreCase(PlatnikEmail, q)
                || ContainsIgnoreCase(StatusNazwa, q)
                || ContainsIgnoreCase(Wystawil, q)
                || MatchesAmount(WartBrutto, q)
                || ContainsIgnoreCase(GlsPreparingBoxParcelNumber, q)
                || ContainsIgnoreCase(GlsPickupParcelNumber, q)
                || ContainsIgnoreCase(GlsShipment?.NrPrzyg, q)
                || ContainsIgnoreCase(GlsShipment?.NrNad, q))
            {
                return true;
            }

            var qDigits = new string(q.Where(char.IsDigit).ToArray());
            if (qDigits.Length >= 3)
            {
                if (NipDigits(KontrahentNip).Contains(qDigits, StringComparison.Ordinal)
                    || NipDigits(PlatnikNip).Contains(qDigits, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NipDigits(string? nip) =>
            new string((nip ?? "").Where(char.IsDigit).ToArray());

        private static bool ContainsIgnoreCase(string? value, string query) =>
            !string.IsNullOrEmpty(value)
            && value.Contains(query, StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Format N2 (pl-PL) używa NBSP jako separatora tysięcy — użytkownik wpisuje zwykłą spację.
        /// Porównujemy też wersję bez spacji („33500” vs „33 500,00”).
        /// </summary>
        private static bool MatchesAmount(decimal amount, string query)
        {
            var display = amount.ToString("N2", CultureInfo.CurrentCulture);
            var normDisplay = NormalizeSpaces(display);
            var normQuery = NormalizeSpaces(query);
            if (normDisplay.Contains(normQuery, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var compactDisplay = CompactNumeric(display);
            var compactQuery = CompactNumeric(query);
            return compactQuery.Length > 0
                && compactDisplay.Contains(compactQuery, StringComparison.Ordinal);
        }

        private static string NormalizeSpaces(string value) =>
            (value ?? "")
                .Replace('\u00A0', ' ')  // NBSP (pl-PL NumberGroupSeparator)
                .Replace('\u202F', ' ')  // narrow no-break space
                .Replace('\u2009', ' '); // thin space

        private static string CompactNumeric(string value) =>
            new string((value ?? "").Where(c => char.IsDigit(c) || c is ',' or '.').ToArray());

        /// <summary>
        /// Adres dostawy z kartoteki kontrahenta (GET /kontrahenci/{id}, pola adr_Dostawa*),
        /// gdy w Subiekcie zaznaczono „Adres dostawy” (kh_AdresDostawy=1).
        /// </summary>
        public bool HasAdresDostawy { get; set; }
        public string AdresDostawyNazwa { get; set; } = "";
        public string AdresDostawyAdres { get; set; } = "";
        public string AdresDostawyUlica { get; set; } = "";
        public string AdresDostawyKodPocztowy { get; set; } = "";
        public string AdresDostawyMiejscowosc { get; set; } = "";
        public string AdresDostawyKrajKod { get; set; } = "";
        public string AdresDostawyTelefon { get; set; } = "";

        public int PlatnikId { get; set; }
        public string PlatnikNazwa { get; set; } = "";
        public string PlatnikNip { get; set; } = "";
        public string PlatnikEmail { get; set; } = "";
        public string PlatnikAdres { get; set; } = "";

        public List<SubiektDocumentLine> Pozycje { get; set; } = new();

        public string WartBruttoDisplay => $"{WartBrutto:N2}";
        public string WartNettoDisplay => $"{WartNetto:N2}";
        public string PrzesylkaDisplay => Przesylka?.ListDisplayText ?? "";
        public string GlsShipmentDisplay => GlsShipment?.ListDisplayText ?? "";

        public bool IsGlsPobranie =>
            KartaNazwa.Equals("GLS pobranie", StringComparison.OrdinalIgnoreCase);

        public string PlatnoscDisplay
        {
            get
            {
                var plat = (PlatNazwa ?? "").Trim();
                var karta = (KartaNazwa ?? "").Trim();
                if (string.IsNullOrEmpty(plat))
                {
                    return karta;
                }

                if (string.IsNullOrEmpty(karta))
                {
                    return plat;
                }

                return $"{plat} · {karta}";
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SubiektPrzesylka
    {
        public string Typ { get; set; } = "";
        public int Id { get; set; }
        /// <summary>Numery listów w przygotowalni GLS (JSON: nr_przyg).</summary>
        public string NrListuPrzygotowalnia { get; set; } = "";
        /// <summary>Numery listów już nadanych (JSON: nr_nad).</summary>
        public string NrListuNadane { get; set; } = "";

        public bool HasId => Id > 0;

        public bool HasAnyNrListu =>
            !string.IsNullOrWhiteSpace(NrListuPrzygotowalnia)
            || !string.IsNullOrWhiteSpace(NrListuNadane);

        public bool IsEmpty => !HasId && !HasAnyNrListu;

        public bool IsPreparingBoxOnly =>
            HasId && string.IsNullOrWhiteSpace(NrListuPrzygotowalnia);

        public bool MatchesGls(int id, string? nrListuPrzygotowalnia, string? nrListuNadane)
        {
            if (IsEmpty || Id != id)
            {
                return false;
            }

            return SameNrListu(NrListuPrzygotowalnia, nrListuPrzygotowalnia)
                && SameNrListu(NrListuNadane, nrListuNadane);
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

        public string ListDisplayText
        {
            get
            {
                var numbers = CombineNrListu(NrListuPrzygotowalnia, NrListuNadane);
                if (!string.IsNullOrEmpty(numbers))
                {
                    return numbers;
                }

                return HasId ? Id.ToString(CultureInfo.InvariantCulture) : "";
            }
        }

        public string SummaryText
        {
            get
            {
                var bits = new List<string>();
                if (HasId)
                {
                    bits.Add($"id {Id}");
                }

                if (!string.IsNullOrWhiteSpace(NrListuPrzygotowalnia))
                {
                    bits.Add($"przygotowalnia {NrListuPrzygotowalnia}");
                }

                if (!string.IsNullOrWhiteSpace(NrListuNadane))
                {
                    bits.Add($"nadane {NrListuNadane}");
                }

                return string.Join(", ", bits);
            }
        }

        public static string CombineNrListu(params string?[] values) =>
            NormalizeNrListu(string.Join(",", values.SelectMany(SplitNrListu)));
    }

    public class SubiektDocumentLine
    {
        public int Plu { get; set; }
        public string Symbol { get; set; } = "";
        public string Nazwa { get; set; } = "";
        public decimal Ilosc { get; set; }
        public decimal CenaNetto { get; set; }
        public decimal CenaBrutto { get; set; }
        public decimal WartoscNetto => Ilosc * CenaNetto;
        public decimal WartoscBrutto => Ilosc * CenaBrutto;
    }

    public class SubiektRelatedDocument
    {
        public int DokId { get; set; }
        public int DokTyp { get; set; }
        public string NrPelny { get; set; } = "";
        public string TypKod => SubiektApiDocumentTypes.FromDokTyp(DokTyp);
        public bool IsFs => DokTyp == SubiektApiDocumentTypes.Fs;
        public bool IsWz => DokTyp == SubiektApiDocumentTypes.Wz;
        public bool IsZk => DokTyp == SubiektApiDocumentTypes.Zk;
    }

    public class SubiektRelatedDocuments
    {
        public SubiektRelatedDocument? Zrodlowy { get; set; }
        public List<SubiektRelatedDocument> Pochodne { get; set; } = new();

        public string FormatDisplayText() =>
            string.Join(Environment.NewLine, ToDisplayList().Select(d => d.NrPelny));

        public List<SubiektRelatedDocument> ToDisplayList()
        {
            var list = new List<SubiektRelatedDocument>();
            AddUnique(list, Zrodlowy);
            foreach (var doc in Pochodne)
            {
                AddUnique(list, doc);
            }

            return list;
        }

        private static void AddUnique(List<SubiektRelatedDocument> list, SubiektRelatedDocument? doc)
        {
            if (doc == null || doc.DokId <= 0)
            {
                return;
            }

            var nr = doc.NrPelny?.Trim() ?? "";
            if (nr.Length == 0)
            {
                return;
            }

            if (list.Any(d => d.DokId == doc.DokId))
            {
                return;
            }

            list.Add(doc);
        }
    }

    /// <summary>
    /// dok_DoDokNrPelny w Subiekcie bywa wypełnione numerem oryginalnym (np. „trasa 27.04”)
    /// zamiast pełnego numeru dokumentu (ZK/WZ/FS…).
    /// </summary>
    public static class SubiektDocumentNumber
    {
        private static readonly string[] Prefixes = { "ZK ", "WZ ", "FS ", "ZD " };

        public static bool IsValid(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var text = value.Trim();
            return Prefixes.Any(p => text.StartsWith(p, StringComparison.OrdinalIgnoreCase));
        }

        public static string Sanitize(string? value) =>
            IsValid(value) ? value!.Trim() : "";
    }

    public static class SubiektApiDocumentTypes
    {
        public const int Fs = 2;
        public const int Wz = 11;
        public const int Zk = 16;

        public static string FromDokTyp(int dokTyp) => dokTyp switch
        {
            Fs => "fs",
            Wz => "wz",
            Zk => "zk",
            _ => "zk"
        };

        public static int FromNrPelny(string? nrPelny)
        {
            var text = (nrPelny ?? "").Trim();
            if (text.StartsWith("FS ", StringComparison.OrdinalIgnoreCase)) return Fs;
            if (text.StartsWith("WZ ", StringComparison.OrdinalIgnoreCase)) return Wz;
            if (text.StartsWith("ZK ", StringComparison.OrdinalIgnoreCase)) return Zk;
            return Zk;
        }

        /// <summary>Dokumenty FS/WZ (related lookup cache GLS).</summary>
        public static bool IsFsOrWz(int dokTyp) => dokTyp is Fs or Wz;

        public static bool IsFsOrWz(string? typKod)
        {
            var typ = (typKod ?? "").Trim().ToLowerInvariant();
            return typ is "fs" or "wz";
        }

        [Obsolete("Use IsFsOrWz")]
        public static bool SupportsPrzesylka(int dokTyp) => IsFsOrWz(dokTyp);

        [Obsolete("Use IsFsOrWz")]
        public static bool SupportsPrzesylka(string? typKod) => IsFsOrWz(typKod);
    }

    public class SubiektDocumentPage
    {
        public List<SubiektDocument> Items { get; set; } = new();
        public int Page { get; set; }
        public int PageSize { get; set; }
        public int TotalCount { get; set; }
        public int TotalPages { get; set; }
    }
}
