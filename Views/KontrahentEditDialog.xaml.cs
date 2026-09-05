using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class KontrahentEditDialog : Window
    {
        private const int SubiektNazwaMaxLength = 50;
        private const int SubiektNrDomuMaxLength = 10;
        private const int SubiektNrLokaluMaxLength = 10;

        private readonly SubiektApiService _api;
        private readonly int _khId;
        private readonly Order? _orderShippingSource;
        private int? _panstwoId;
        private int? _wojewodztwoId;
        private int? _dostawaPanstwoId;
        private int? _dostawaWojewodztwoId;
        private bool _isSaving;
        private bool _isRefreshing;

        public KontrahentDetails? SavedDetails { get; private set; }

        public KontrahentEditDialog(KontrahentDetails details, Order? orderShippingSource = null, SubiektApiService? api = null)
        {
            InitializeComponent();
            _api = api ?? new SubiektApiService();
            _khId = details.Id;
            _orderShippingSource = orderShippingSource;
            ApplyKontrahentIds(details);

            HeaderTitleText.Text = GetCompanyDisplayName(details);
            if (string.IsNullOrWhiteSpace(HeaderTitleText.Text))
                HeaderTitleText.Text = "Adres wysyłki";

            BindReadOnlyCard(details);
            PrefillShippingFromOrder(orderShippingSource);
            BindCurrentDostawaCard(details, orderShippingSource);

            DostawaAktywnyCheck.IsChecked = false;
            UpdateDostawaFieldsEnabled();
        }

        private void ApplyKontrahentIds(KontrahentDetails details)
        {
            _panstwoId = details.PanstwoId;
            _wojewodztwoId = details.WojewodztwoId;
            _dostawaPanstwoId = details.DostawaPanstwoId;
            _dostawaWojewodztwoId = details.DostawaWojewodztwoId;
        }

        private async void OpenSubiektKontrahentButton_Click(object sender, RoutedEventArgs e)
        {
            if (_khId <= 0)
            {
                MessageBox.Show(
                    "Brak identyfikatora kontrahenta (kh_Id).",
                    "Kontrahent",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            try
            {
                if (DostawaAktywnyCheck.IsChecked == true
                    && !TryValidateDostawaFields(out var validationError))
                {
                    MessageBox.Show(
                        validationError,
                        "Adres wysyłki",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                AdresDostawyPrefill? prefill = null;
                if (DostawaAktywnyCheck.IsChecked == true)
                {
                    prefill = new AdresDostawyPrefill
                    {
                        Nazwa = NullIfEmpty(DostawaNazwaBox.Text),
                        Ulica = NullIfEmpty(DostawaUlicaBox.Text),
                        NrDomu = NullIfEmpty(DostawaNrDomuBox.Text),
                        NrLokalu = NullIfEmpty(DostawaNrLokaluBox.Text),
                        Kod = NullIfEmpty(DostawaKodBox.Text),
                        Miejscowosc = NullIfEmpty(DostawaMiejscowoscBox.Text),
                        PanstwoId = _dostawaPanstwoId ?? _panstwoId,
                        WojewodztwoId = _dostawaWojewodztwoId ?? _wojewodztwoId
                    };
                    Debug($"Otwieranie kartoteki kh_Id={_khId} z prefill adresu dostawy (Sfera)", "KontrahentEditDialog");
                }
                else
                {
                    Debug($"Otwieranie kartoteki kontrahenta kh_Id={_khId} przez Sferę", "KontrahentEditDialog");
                }

                Mouse.OverrideCursor = Cursors.Wait;
                var subiektService = new SubiektService();
                // Wyswietl() blokuje do zamknięcia okna Subiekta (musi być na wątku UI / STA)
                subiektService.OtworzKartotekeKontrahenta(_khId, prefill);
            }
            catch (Exception ex)
            {
                Error(ex, "KontrahentEditDialog", "Błąd otwierania kartoteki w Subiekcie");
                MessageBox.Show(
                    $"Nie udało się otworzyć kartoteki kontrahenta w Subiekcie:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }

            // Po zamknięciu Subiekta — odśwież kartę porównania z zamówieniem
            await RefreshDostawaComparisonAsync();
        }

        private static string? NullIfEmpty(string? value) =>
            string.IsNullOrWhiteSpace(value) ? null : value.Trim();

        private static string GetCompanyDisplayName(KontrahentDetails d)
        {
            if (!string.IsNullOrWhiteSpace(d.NazwaPelna))
                return d.NazwaPelna.Trim();
            if (!string.IsNullOrWhiteSpace(d.Nazwa))
                return d.Nazwa.Trim();
            return d.Symbol?.Trim() ?? "";
        }

        private static string? GetPersonDisplayName(KontrahentDetails d)
        {
            var nazwa = d.Nazwa?.Trim() ?? "";
            var firma = d.NazwaPelna?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(nazwa))
                return null;
            if (string.IsNullOrWhiteSpace(firma))
                return null;
            if (Order.NamesMatch(nazwa, firma))
                return null;
            return nazwa;
        }

        private void BindReadOnlyCard(KontrahentDetails d)
        {
            var firma = d.NazwaPelna?.Trim() ?? "";
            var nazwa = d.Nazwa?.Trim() ?? "";
            var osoba = GetPersonDisplayName(d);

            CardOsobaText.Text = osoba ?? "";
            CardFirmaText.Text = !string.IsNullOrWhiteSpace(firma)
                ? firma
                : (string.IsNullOrWhiteSpace(osoba) ? nazwa : "");

            CardNipText.Text = d.Nip?.Trim() ?? "";
            CardEmailText.Text = d.Email?.Trim() ?? "";

            var adresParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(d.Adres)) adresParts.Add(d.Adres.Trim());
            else if (!string.IsNullOrWhiteSpace(d.Ulica))
            {
                var street = d.Ulica.Trim();
                var nr = $"{d.NrDomu} {d.NrLokalu}".Trim();
                adresParts.Add(string.IsNullOrWhiteSpace(nr) ? street : $"{street} {nr}");
            }
            var city = $"{d.Kod} {d.Miejscowosc}".Trim();
            if (!string.IsNullOrWhiteSpace(city)) adresParts.Add(city);
            if (!string.IsNullOrWhiteSpace(d.Panstwo) && !d.Panstwo.Equals("Polska", StringComparison.OrdinalIgnoreCase))
                adresParts.Add(d.Panstwo.Trim());
            CardAdresText.Text = string.Join(", ", adresParts);
        }

        private static string? GetDostawaPersonDisplayName(KontrahentDetails d)
        {
            var nazwa = d.DostawaNazwa?.Trim() ?? "";
            var firma = d.DostawaNazwaPelna?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(nazwa))
                return null;
            if (string.IsNullOrWhiteSpace(firma))
                return null;
            if (Order.NamesMatch(nazwa, firma))
                return null;
            return nazwa;
        }

        private void BindCurrentDostawaCard(KontrahentDetails d, Order? order)
        {
            if (!d.HasAdresDostawy)
            {
                CurrentDostawaFieldsPanel.Visibility = Visibility.Collapsed;
                CurrentDostawaEmptyText.Visibility = Visibility.Visible;
                CurrentDostawaCard.BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0));
                CurrentDostawaCard.Background = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA));
                CurrentDostawaTitle.Foreground = new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75));
                CurrentDostawaCompareText.Text = "";
                CurrentDostawaDiffText.Visibility = Visibility.Collapsed;
                return;
            }

            CurrentDostawaFieldsPanel.Visibility = Visibility.Visible;
            CurrentDostawaEmptyText.Visibility = Visibility.Collapsed;

            var firma = d.DostawaNazwaPelna?.Trim() ?? "";
            var nazwa = d.DostawaNazwa?.Trim() ?? "";
            var osoba = GetDostawaPersonDisplayName(d);

            CurrentDostawaOsobaText.Text = osoba ?? "";
            CurrentDostawaFirmaText.Text = !string.IsNullOrWhiteSpace(firma)
                ? firma
                : (string.IsNullOrWhiteSpace(osoba) ? nazwa : "");
            CurrentDostawaNipText.Text = d.DostawaNip?.Trim() ?? "";
            CurrentDostawaTelefonText.Text = d.DostawaTelefon?.Trim() ?? "";

            var adresParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(d.DostawaAdres)) adresParts.Add(d.DostawaAdres.Trim());
            else if (!string.IsNullOrWhiteSpace(d.DostawaUlica))
            {
                var street = d.DostawaUlica.Trim();
                var nr = $"{d.DostawaNrDomu} {d.DostawaNrLokalu}".Trim();
                adresParts.Add(string.IsNullOrWhiteSpace(nr) ? street : $"{street} {nr}");
            }
            var cityLine = $"{d.DostawaKod} {d.DostawaMiejscowosc}".Trim();
            if (!string.IsNullOrWhiteSpace(cityLine)) adresParts.Add(cityLine);
            if (!string.IsNullOrWhiteSpace(d.DostawaPanstwo)
                && !d.DostawaPanstwo.Equals("Polska", StringComparison.OrdinalIgnoreCase))
                adresParts.Add(d.DostawaPanstwo.Trim());
            CurrentDostawaAdresText.Text = string.Join(", ", adresParts);

            if (order == null || !order.HasShippingAddress)
            {
                CurrentDostawaCard.BorderBrush = new SolidColorBrush(Color.FromRgb(0x90, 0xA4, 0xAE));
                CurrentDostawaCard.Background = new SolidColorBrush(Color.FromRgb(0xEC, 0xEF, 0xF1));
                CurrentDostawaTitle.Foreground = new SolidColorBrush(Color.FromRgb(0x45, 0x5A, 0x64));
                CurrentDostawaCompareText.Text = "Brak adresu wysyłki w zamówieniu do porównania.";
                CurrentDostawaCompareText.Foreground = new SolidColorBrush(Color.FromRgb(0x54, 0x6E, 0x7A));
                CurrentDostawaDiffText.Visibility = Visibility.Collapsed;
                return;
            }

            var diffs = CompareDostawaWithOrderShipping(d, order);
            if (diffs.Count == 0)
            {
                CurrentDostawaCard.BorderBrush = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
                CurrentDostawaCard.Background = new SolidColorBrush(Color.FromRgb(0xE8, 0xF5, 0xE9));
                CurrentDostawaTitle.Foreground = new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32));
                CurrentDostawaCompareText.Text = "Taki sam jak adres wysyłki z zamówienia.";
                CurrentDostawaCompareText.Foreground = new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32));
                CurrentDostawaDiffText.Visibility = Visibility.Collapsed;
            }
            else if (IsNameOnlyDiff(diffs))
            {
                // Jak w szczegółach zamówienia: tylko firma/nazwa inna, adres ten sam → warning żółty
                CurrentDostawaCard.BorderBrush = new SolidColorBrush(Color.FromRgb(0xF9, 0xA8, 0x25));
                CurrentDostawaCard.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF8, 0xE1));
                CurrentDostawaTitle.Foreground = new SolidColorBrush(Color.FromRgb(0xF5, 0x7F, 0x17));
                CurrentDostawaCompareText.Text = "Inna nazwa firmy ale adres ten sam.";
                CurrentDostawaCompareText.Foreground = new SolidColorBrush(Color.FromRgb(0xF5, 0x7F, 0x17));
                CurrentDostawaDiffText.Visibility = Visibility.Collapsed;
            }
            else
            {
                CurrentDostawaCard.BorderBrush = new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36));
                CurrentDostawaCard.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE));
                CurrentDostawaTitle.Foreground = new SolidColorBrush(Color.FromRgb(0xC6, 0x28, 0x28));
                CurrentDostawaCompareText.Text = "Różni się od adresu wysyłki z zamówienia.";
                CurrentDostawaCompareText.Foreground = new SolidColorBrush(Color.FromRgb(0xC6, 0x28, 0x28));
                CurrentDostawaDiffText.Text = "Różnice: " + string.Join(", ", diffs);
                CurrentDostawaDiffText.Foreground = new SolidColorBrush(Color.FromRgb(0xC6, 0x28, 0x28));
                CurrentDostawaDiffText.Visibility = Visibility.Visible;
            }
        }

        private static bool IsNameOnlyDiff(List<string> diffs) =>
            diffs.Count == 1 && diffs[0] == "nazwa";

        private static List<string> CompareDostawaWithOrderShipping(KontrahentDetails d, Order order)
        {
            var diffs = new List<string>();

            var orderName = !string.IsNullOrWhiteSpace(order.ShippingCompanyDisplay)
                ? order.ShippingCompanyDisplay
                : order.ShippingDisplayName;

            // Zgodność z Nazwa albo NazwaPelna (Subiekt bywa: Nazwa = osoba, Pelna = firma)
            if (!Order.NamesMatch(orderName, d.DostawaNazwa)
                && !Order.NamesMatch(orderName, d.DostawaNazwaPelna))
                diffs.Add("nazwa");

            var orderParsed = ParseStreetParts(order.ShippingAddress1, order.ShippingAddress2);
            var dostawaParsed = ParseDostawaStreetParts(d);

            if (!StreetNamesMatch(orderParsed.Ulica, dostawaParsed.Ulica))
                diffs.Add("ulica");

            if (!string.Equals(Normalize(orderParsed.NrDomu), Normalize(dostawaParsed.NrDomu), StringComparison.Ordinal))
                diffs.Add("nr domu");

            if (!string.Equals(Normalize(EffectiveNrLokalu(orderParsed.NrLokalu, orderParsed.Ulica)),
                    Normalize(EffectiveNrLokalu(dostawaParsed.NrLokalu, dostawaParsed.Ulica)), StringComparison.Ordinal))
                diffs.Add("nr lokalu");

            if (!string.Equals(Normalize(order.ShippingPostcode), Normalize(d.DostawaKod), StringComparison.Ordinal))
                diffs.Add("kod pocztowy");

            if (!string.Equals(Normalize(order.ShippingCity), Normalize(d.DostawaMiejscowosc), StringComparison.Ordinal))
                diffs.Add("miasto");

            return diffs;
        }

        /// <summary>
        /// Pola dostawy z Subiekta — nr domu/lokal często tylko w adr_DostawaAdres.
        /// </summary>
        private static (string Ulica, string NrDomu, string NrLokalu) ParseDostawaStreetParts(KontrahentDetails d)
        {
            var ulica = CleanUlica(d.DostawaUlica?.Trim() ?? "");
            var nrDomu = d.DostawaNrDomu?.Trim() ?? "";
            var nrLokalu = d.DostawaNrLokalu?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(nrDomu) && !string.IsNullOrWhiteSpace(d.DostawaAdres))
            {
                var fromAdres = ParseStreetParts(d.DostawaAdres, null);
                if (string.IsNullOrWhiteSpace(ulica))
                    ulica = fromAdres.Ulica;
                nrDomu = fromAdres.NrDomu;
                if (string.IsNullOrWhiteSpace(nrLokalu))
                    nrLokalu = fromAdres.NrLokalu;
            }

            return (CleanUlica(ulica), nrDomu, EffectiveNrLokalu(nrLokalu, ulica));
        }

        private static bool StreetNamesMatch(string? a, string? b)
        {
            var na = NormalizeStreetTokens(Normalize(a));
            var nb = NormalizeStreetTokens(Normalize(b));
            if (string.Equals(na, nb, StringComparison.Ordinal))
                return true;
            return string.IsNullOrEmpty(na) && string.IsNullOrEmpty(nb);
        }

        private static string NormalizeStreetTokens(string value)
        {
            var tokens = value.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(t => t is not "ul" and not "ul." and not "ulicy" and not "al" and not "al.")
                .ToArray();
            return CleanUlica(string.Join(" ", tokens));
        }

        private static string CleanUlica(string ulica)
        {
            return ulica.Trim().TrimEnd(',', '.', ' ').Trim();
        }

        /// <summary>OpenCart czasem powtarza nazwę ulicy w address_2 — ignoruj jako lokal.</summary>
        private static bool IsDuplicateStreetReference(string? value, string ulica)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            return StreetNamesMatch(CleanUlica(ulica), value.Trim());
        }

        private static string EffectiveNrLokalu(string? nrLokalu, string ulica)
        {
            return IsDuplicateStreetReference(nrLokalu, ulica) ? "" : (nrLokalu?.Trim() ?? "");
        }

        private static string Normalize(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";
            var decoded = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(value)).Trim();
            return string.Join(" ", decoded.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                .ToLowerInvariant();
        }

        private void PrefillShippingFromOrder(Order? order)
        {
            if (order == null)
                return;

            DostawaNazwaBox.Text = !string.IsNullOrWhiteSpace(order.ShippingCompany)
                ? order.ShippingCompany.Trim()
                : order.ShippingDisplayName;

            var parsed = ParseStreetParts(order.ShippingAddress1, order.ShippingAddress2);
            DostawaUlicaBox.Text = parsed.Ulica;
            DostawaNrDomuBox.Text = parsed.NrDomu;
            DostawaNrLokaluBox.Text = parsed.NrLokalu;
            DostawaKodBox.Text = order.ShippingPostcode?.Trim() ?? "";
            DostawaMiejscowoscBox.Text = order.ShippingCity?.Trim() ?? "";
        }

        /// <summary>
        /// OpenCart: address_1 = ulica (+ czasem numer), address_2 = nr domu albo nr lokalu / dopisek.
        /// </summary>
        private static (string Ulica, string NrDomu, string NrLokalu) ParseStreetParts(string? address1, string? address2)
        {
            var a1 = (address1 ?? "").Trim();
            var a2 = (address2 ?? "").Trim();

            string ulica = a1;
            string nrDomu = "";
            string nrLokalu = "";

            // address_1: „ul. Wojskowa 3 / L4” lub „Kwiatowa 16”
            var fromA1 = TrySplitStreetAndNumbers(a1);
            if (fromA1 != null)
            {
                ulica = fromA1.Value.Ulica;
                nrDomu = fromA1.Value.NrDomu;
                nrLokalu = fromA1.Value.NrLokalu;
            }

            if (string.IsNullOrWhiteSpace(a2))
                return (CleanUlica(ulica), nrDomu, EffectiveNrLokalu(nrLokalu, ulica));

            if (IsDuplicateStreetReference(a2, ulica))
                return (CleanUlica(ulica), nrDomu, EffectiveNrLokalu(nrLokalu, ulica));

            // address_2: „16”, „16A”, „16 / 2”, „16/L4”, „L4” — nie nazwa wsi bez numeru
            var fromA2 = TryParseHouseOrLocal(a2);
            if (fromA2 != null && Order.LooksLikeHouseOrUnitPart(a2))
            {
                if (!string.IsNullOrWhiteSpace(fromA2.Value.NrDomu) && string.IsNullOrWhiteSpace(nrDomu))
                    nrDomu = fromA2.Value.NrDomu;
                else if (!string.IsNullOrWhiteSpace(fromA2.Value.NrDomu) && !string.IsNullOrWhiteSpace(nrDomu)
                         && string.IsNullOrWhiteSpace(nrLokalu) && string.IsNullOrWhiteSpace(fromA2.Value.NrLokalu))
                    // nr domu już jest — samotny numer w address_2 traktuj jako lokal tylko gdy wygląda na lokal
                    nrLokalu = fromA2.Value.NrDomu;

                if (!string.IsNullOrWhiteSpace(fromA2.Value.NrLokalu) && string.IsNullOrWhiteSpace(nrLokalu))
                    nrLokalu = fromA2.Value.NrLokalu;
            }
            else if (string.IsNullOrWhiteSpace(nrLokalu) && !IsDuplicateStreetReference(a2, ulica)
                     && Order.LooksLikeHouseOrUnitPart(a2))
            {
                nrLokalu = a2;
            }

            return (CleanUlica(ulica), nrDomu, EffectiveNrLokalu(nrLokalu, ulica));
        }

        private static (string Ulica, string NrDomu, string NrLokalu)? TrySplitStreetAndNumbers(string address1)
        {
            if (string.IsNullOrWhiteSpace(address1))
                return null;

            // „… 3 / L4”, „… 3/L4”, „… 16”, „Nowy Świat, 22A/Nowy Świat”
            var m = System.Text.RegularExpressions.Regex.Match(
                address1,
                @"^(?<ulica>.+?)\s+(?<dom>\d+[A-Za-z]?)\s*(?:[/\\]\s*(?<lok>.+))?\s*$");
            if (!m.Success)
                return null;

            return (
                CleanUlica(m.Groups["ulica"].Value),
                m.Groups["dom"].Value.Trim(),
                m.Groups["lok"].Success ? m.Groups["lok"].Value.Trim() : "");
        }

        private static (string NrDomu, string NrLokalu)? TryParseHouseOrLocal(string address2)
        {
            if (string.IsNullOrWhiteSpace(address2))
                return null;

            // „16 / 2”, „16/L4”
            var split = System.Text.RegularExpressions.Regex.Match(
                address2,
                @"^(?<dom>\d+[A-Za-z]?)\s*[/\\]\s*(?<lok>.+)$");
            if (split.Success)
                return (split.Groups["dom"].Value.Trim(), split.Groups["lok"].Value.Trim());

            // Sam nr domu: „16”, „16A”
            if (System.Text.RegularExpressions.Regex.IsMatch(address2, @"^\d+[A-Za-z]?$"))
                return (address2.Trim(), "");

            // Lokal / dopisek bez nr domu: „L4”, „m.5”
            return ("", address2.Trim());
        }

        private void DostawaAktywnyCheck_Changed(object sender, RoutedEventArgs e)
        {
            UpdateDostawaFieldsEnabled();
        }

        private void UpdateDostawaFieldsEnabled()
        {
            var active = DostawaAktywnyCheck.IsChecked == true;
            DostawaFieldsPanel.IsEnabled = active;
            DostawaFieldsPanel.Opacity = active ? 1 : 0.65;
            SaveButton.IsEnabled = active;
            OpenSubiektKontrahentButton.ToolTip = active
                ? "Otwórz kartotekę w Subiekcie z wypełnionym adresem wysyłki — potem zapisz w Subiekcie"
                : "Otwórz kartotekę kontrahenta w Subiekcie GT (Sfera)";
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isSaving || _khId <= 0 || DostawaAktywnyCheck.IsChecked != true)
                return;

            if (!TryValidateDostawaFields(out var validationError))
            {
                MessageBox.Show(
                    validationError,
                    "Adres wysyłki",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _isSaving = true;
            SaveButton.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                var request = BuildRequest();
                var saved = await _api.UpdateKontrahentAsync(_khId, request);
                if (saved == null)
                {
                    MessageBox.Show(
                        "API nie zwróciło zaktualizowanych danych kontrahenta.",
                        "Adres wysyłki",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                SavedDetails = saved;
                ApplyKontrahentIds(saved);
                BindReadOnlyCard(saved);
                BindCurrentDostawaCard(saved, _orderShippingSource);
                DostawaAktywnyCheck.IsChecked = false;
                UpdateDostawaFieldsEnabled();
                Debug($"Po zapisie API odświeżono porównanie adresu dostawy kh_Id={_khId}", "KontrahentEditDialog");
            }
            catch (Exception ex)
            {
                Error(ex, "KontrahentEditDialog", "Błąd zapisu adresu wysyłki");
                MessageBox.Show(
                    $"Nie udało się zapisać adresu wysyłki:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _isSaving = false;
                SaveButton.IsEnabled = DostawaAktywnyCheck.IsChecked == true;
                Mouse.OverrideCursor = null;
            }
        }

        /// <summary>
        /// Pobiera aktualną kartotekę z API i odświeża kartę „Adres wysyłki w Subiekcie”
        /// oraz porównanie z adresem z zamówienia.
        /// </summary>
        private async System.Threading.Tasks.Task RefreshDostawaComparisonAsync()
        {
            if (_khId <= 0 || _isRefreshing)
                return;

            _isRefreshing = true;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                var details = await _api.GetKontrahentDetailsAsync(_khId);
                if (details == null)
                {
                    Warning($"Nie udało się odświeżyć kontrahenta kh_Id={_khId} po edycji", "KontrahentEditDialog");
                    return;
                }

                SavedDetails = details;
                ApplyKontrahentIds(details);
                BindReadOnlyCard(details);
                BindCurrentDostawaCard(details, _orderShippingSource);
                DostawaAktywnyCheck.IsChecked = false;
                UpdateDostawaFieldsEnabled();
                Debug($"Odświeżono porównanie adresu dostawy kh_Id={_khId}", "KontrahentEditDialog");
            }
            catch (Exception ex)
            {
                Error(ex, "KontrahentEditDialog", "Błąd odświeżania porównania adresu dostawy");
            }
            finally
            {
                _isRefreshing = false;
                Mouse.OverrideCursor = null;
            }
        }

        private KontrahentUpdateRequest BuildRequest()
        {
            static string? NullIfEmpty(string? value) =>
                string.IsNullOrWhiteSpace(value) ? null : value.Trim();

            return new KontrahentUpdateRequest
            {
                AdresDostawy = new KontrahentSecondaryAddressUpdate
                {
                    Aktywny = true,
                    Nazwa = NullIfEmpty(DostawaNazwaBox.Text),
                    Ulica = NullIfEmpty(DostawaUlicaBox.Text),
                    NrDomu = NullIfEmpty(DostawaNrDomuBox.Text),
                    NrLokalu = NullIfEmpty(DostawaNrLokaluBox.Text),
                    Kod = NullIfEmpty(DostawaKodBox.Text),
                    Miejscowosc = NullIfEmpty(DostawaMiejscowoscBox.Text),
                    PanstwoId = _dostawaPanstwoId ?? _panstwoId,
                    WojewodztwoId = _dostawaWojewodztwoId ?? _wojewodztwoId
                }
            };
        }

        private bool TryValidateDostawaFields(out string errorMessage)
        {
            var nazwa = DostawaNazwaBox.Text?.Trim() ?? "";
            if (nazwa.Length > SubiektNazwaMaxLength)
            {
                errorMessage =
                    $"Nazwa / odbiorca może mieć maksymalnie {SubiektNazwaMaxLength} znaków (limit Subiekta).\n\n" +
                    $"Obecna wartość ma {nazwa.Length} znaków:\n\"{nazwa}\"\n\n" +
                    "Skróć pole „Nazwa / odbiorca” przed zapisem lub otwarciem Subiekta.";
                return false;
            }

            var nrDomu = DostawaNrDomuBox.Text?.Trim() ?? "";
            if (nrDomu.Length > SubiektNrDomuMaxLength)
            {
                errorMessage =
                    $"Nr domu może mieć maksymalnie {SubiektNrDomuMaxLength} znaków (limit Subiekta).\n\n" +
                    $"Obecna wartość ma {nrDomu.Length} znaków:\n\"{nrDomu}\"\n\n" +
                    "Skróć pole „Nr domu” przed zapisem.";
                return false;
            }

            var nrLokalu = DostawaNrLokaluBox.Text?.Trim() ?? "";
            if (nrLokalu.Length > SubiektNrLokaluMaxLength)
            {
                errorMessage =
                    $"Nr lokalu może mieć maksymalnie {SubiektNrLokaluMaxLength} znaków (limit Subiekta).\n\n" +
                    $"Obecna wartość ma {nrLokalu.Length} znaków:\n\"{nrLokalu}\"\n\n" +
                    "Skróć pole „Nr lokalu / dopisek” przed zapisem.";
                return false;
            }

            errorMessage = "";
            return true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
