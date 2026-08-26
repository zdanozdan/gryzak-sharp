using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.Services;

namespace Gryzak.Views
{
    public partial class GlsParcelSearchDialog : Window
    {
        private readonly ConfigService _configService;

        public GlsParcelSearchDialog(ConfigService configService)
        {
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            InitializeComponent();
            Loaded += (_, _) => ParcelNumberTextBox.Focus();
            StatusTextBlock.Text = "Podaj numer paczki i kliknij „Szukaj”.";
        }

        private async void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            var input = (ParcelNumberTextBox.Text ?? "").Trim();
            var targetDigits = NormalizeDigits(input);

            ResultTextBox.Clear();
            StatusTextBlock.Text = "";

            if (string.IsNullOrWhiteSpace(targetDigits))
            {
                StatusTextBlock.Text = "Nieprawidłowy numer paczki (wpisz cyfry).";
                return;
            }

            SearchButton.IsEnabled = false;
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            try
            {
                var config = _configService.LoadGlsConfig();
                if (string.IsNullOrWhiteSpace(config.UserName) || string.IsNullOrWhiteSpace(config.Password))
                {
                    StatusTextBlock.Text = "Brak loginu/hasła GLS w ustawieniach.";
                    MessageBox.Show(
                        "Uzupełnij login i hasło GLS w menu Ustawienia → Ustawienia GLS.",
                        "Brak konfiguracji GLS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                using var gls = new GlsService(config);
                var sb = new StringBuilder();

                sb.AppendLine($"Środowisko: {config.GetEnvironmentName()}");
                sb.AppendLine($"Numer wejściowy: {input}");
                sb.AppendLine();

                // 1) Bezpośrednie wyszukiwanie po numerze (SOAP)
                sb.AppendLine("1) SOAP: adePickup_ParcelNumberSearch");
                var soap = await gls.SearchParcelNumberAsync(targetDigits);
                AppendShipmentResult(sb, soap);

                if (soap.Success && soap.Consignment != null)
                {
                    ResultTextBox.Text = sb.ToString();
                    StatusTextBlock.Text = "OK (znalezione po SOAP).";
                    return;
                }

                // 2) Fallback: skan przygotowalni/nadań i dopasowanie po numerze
                sb.AppendLine();
                sb.AppendLine("2) Fallback: skan list GLS (przygotowalnia/nadania)");

                var lists = await gls.GetStatusListsAsync();
                if (!lists.Success)
                {
                    sb.AppendLine($"Nie udało się pobrać list GLS: {lists.ErrorMessage}");
                    ResultTextBox.Text = sb.ToString();
                    StatusTextBlock.Text = "Błąd pobierania list.";
                    return;
                }

                var preparing = lists.PreparingBoxItems.FirstOrDefault(i => ParcelNumberMatches(i.ParcelNumber, targetDigits));
                var pickup = lists.PickupItems.FirstOrDefault(i => ParcelNumberMatches(i.ParcelNumber, targetDigits));

                if (preparing is null && pickup is null)
                {
                    sb.AppendLine($"Nie znaleziono paczki {targetDigits} w przygotowalni ani w nadaniach (po numerze).");
                    ResultTextBox.Text = sb.ToString();
                    StatusTextBlock.Text = "Nie znaleziono.";
                    return;
                }

                if (preparing != null)
                {
                    sb.AppendLine($"Znaleziono w przygotowalni: item.Id={preparing.Id}, GLS-NrListu={preparing.ParcelNumber}");
                    var consignment = await gls.GetShipmentAsync(preparing.Id);
                    AppendShipmentResult(sb, consignment);
                }
                else if (pickup != null)
                {
                    sb.AppendLine($"Znaleziono w nadaniach: item.Id={pickup.Id}, GLS-NrListu={pickup.ParcelNumber}");
                    var consignment = await gls.GetPickupShipmentAsync(pickup.Id);
                    AppendShipmentResult(sb, consignment);
                }

                ResultTextBox.Text = sb.ToString();
                StatusTextBlock.Text = "Zakończono.";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Błąd testu.";
                ResultTextBox.Text = ex.ToString();
            }
            finally
            {
                SearchButton.IsEnabled = true;
                Mouse.OverrideCursor = null;
            }
        }

        private void AppendShipmentResult(StringBuilder sb, GlsGetShipmentResult result)
        {
            if (result == null)
            {
                sb.AppendLine("Brak wyniku.");
                return;
            }

            sb.AppendLine(result.Success
                ? "Wynik: Success"
                : $"Wynik: {(result.NotFound ? "NotFound" : "Failed")}");

            if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
            {
                sb.AppendLine("Błąd: " + result.ErrorMessage);
            }

            if (result.Success && result.Consignment != null)
            {
                var c = result.Consignment;
                sb.AppendLine($"GLS id (ExistingId): {(c.ExistingId?.ToString() ?? "-")}");
                sb.AppendLine($"Nadawca/odbiorca (Name1): {c.Name1}");
                if (!string.IsNullOrWhiteSpace(c.Name2))
                {
                    sb.AppendLine($"Name2: {c.Name2}");
                }

                if (!string.IsNullOrWhiteSpace(c.Name3))
                {
                    sb.AppendLine($"Name3: {c.Name3}");
                }

                sb.AppendLine($"Adres: {c.SummaryAddress}");
                sb.AppendLine($"Referencje: {c.References}");
                sb.AppendLine($"COD: {(c.CashOnDelivery ? c.CodAmountDisplay + " zł" : "nie")}");
                sb.AppendLine($"Paczki: {c.GetParcelNumbersDisplay()}");
            }
        }

        private static string NormalizeDigits(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "";
            }

            return Regex.Replace(value, @"\D+", "");
        }

        private static bool ParcelNumberMatches(string? candidate, string targetDigits)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(targetDigits))
            {
                return false;
            }

            // candidate może zawierać np. "123, 456" (lista paczek w konsygnacji)
            var digits = Regex.Matches(candidate, @"\d+")
                .Select(m => m.Value)
                .Where(v => !string.IsNullOrWhiteSpace(v));

            return digits.Any(d => string.Equals(d, targetDigits, StringComparison.OrdinalIgnoreCase));
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

