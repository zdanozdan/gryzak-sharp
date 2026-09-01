using System;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class GlsWaybillDialog : Window
    {
        private readonly GlsConfig _glsConfig;
        private readonly ConfigService _configService;
        private string _existingNrListu;
        private bool _hasIssuedLabel;
        private bool _isBusy;

        public GlsConsignment Consignment { get; private set; }
        public bool DeleteRequested { get; private set; }
        /// <summary>True, gdy w tej sesji okna pobrano etykietę (numery listu nadane przez GLS).</summary>
        public bool LabelIssued { get; private set; }
        /// <summary>Numery paczek odczytane z GLS po etykiecie (CSV / display).</summary>
        public string IssuedParcelNumbers { get; private set; } = "";

        public GlsWaybillDialog(
            GlsConsignment consignment,
            GlsConfig glsConfig,
            string? existingNrListu = null,
            ConfigService? configService = null)
        {
            InitializeComponent();
            _glsConfig = glsConfig ?? new GlsConfig();
            _configService = configService ?? new ConfigService();
            _existingNrListu = (existingNrListu ?? "").Trim();
            Consignment = consignment?.Clone() ?? new GlsConsignment();
            Consignment.NormalizeLimits();
            _hasIssuedLabel = HasIssuedLabel(Consignment, _existingNrListu);

            var isEdit = Consignment.IsEdit;
            var environmentName = _glsConfig.GetEnvironmentName();
            Title = isEdit ? "Edycja w przygotowalni GLS" : "Przygotowalnia GLS";
            TitleText.Text = isEdit ? "Edycja przesyłki w przygotowalni GLS" : "Nowa przesyłka w przygotowalni GLS";
            SubmitButton.Content = isEdit ? "Zapisz" : "Dodaj";
            DeleteButton.Visibility = isEdit ? Visibility.Visible : Visibility.Collapsed;
            LabelButton.Visibility = isEdit ? Visibility.Visible : Visibility.Collapsed;
            ApplyIssuedLabelLock();

            UpdateSubtitle(environmentName);
            Name1TextBox.Text = Consignment.Name1;
            Name2TextBox.Text = Consignment.Name2;
            Name3TextBox.Text = Consignment.Name3;
            StreetTextBox.Text = Consignment.Street;
            CountryTextBox.Text = string.IsNullOrWhiteSpace(Consignment.Country) ? "PL" : Consignment.Country;
            ZipCodeTextBox.Text = Consignment.ZipCode;
            CityTextBox.Text = Consignment.City;
            PhoneTextBox.Text = Consignment.Phone;
            ContactTextBox.Text = Consignment.Contact;
            ReferencesTextBox.Text = Consignment.References;
            NotesTextBox.Text = Consignment.Notes;
            ParcelCountTextBox.Text = Consignment.ParcelCount.ToString(CultureInfo.InvariantCulture);
            WeightTextBox.Text = Consignment.ParcelWeight;
            CashOnDeliveryCheckBox.IsChecked = Consignment.CashOnDelivery;
            UpdateCodAmountHint();
        }

        private void UpdateSubtitle(string? environmentName = null)
        {
            environmentName ??= _glsConfig.GetEnvironmentName();
            var env = string.IsNullOrWhiteSpace(environmentName) ? "" : $"Środowisko: {environmentName}";
            string idInfo;
            if (!Consignment.IsEdit)
            {
                idInfo = "Przesyłka trafi do przygotowalni GLS.";
            }
            else if (!string.IsNullOrWhiteSpace(_existingNrListu))
            {
                idInfo = $"Id w przygotowalni: {Consignment.ExistingId}  •  Nr listu: {_existingNrListu}";
            }
            else
            {
                idInfo = $"Id w przygotowalni: {Consignment.ExistingId}";
            }

            SubtitleText.Text = string.IsNullOrWhiteSpace(env) ? idInfo : $"{env}  •  {idInfo}";
        }

        private void CashOnDeliveryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateCodAmountHint();
        }

        private void UpdateCodAmountHint()
        {
            var enabled = CashOnDeliveryCheckBox.IsChecked == true;
            if (!enabled)
            {
                CodAmountHintText.Visibility = Visibility.Collapsed;
                return;
            }

            CodAmountHintText.Visibility = Visibility.Visible;
            CodAmountHintText.Text = Consignment.CodAmount > 0
                ? $"Kwota pobrania: {Consignment.CodAmountDisplay} zł."
                : "Brak kwoty pobrania — dokument musi mieć wartość brutto większą od zera.";
        }

        private async void LabelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            if (Consignment.ExistingId is not > 0)
            {
                MessageBox.Show(
                    "Etykietę można pobrać dopiero po dodaniu przesyłki do przygotowalni GLS.",
                    "Brak przesyłki",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            _isBusy = true;
            LabelButton.IsEnabled = false;
            try
            {
                var dialog = new GlsLabelDialog(Consignment.ExistingId.Value, _glsConfig, _configService)
                {
                    Owner = this
                };
                dialog.ShowDialog();
                if (!dialog.LabelsLoaded)
                {
                    return;
                }

                LabelIssued = true;
                await RefreshParcelNumbersAfterLabelAsync(Consignment.ExistingId.Value);
                _hasIssuedLabel = true;
                ApplyIssuedLabelLock();
                UpdateSubtitle();
            }
            catch (Exception ex)
            {
                Error(ex, "GlsWaybillDialog", "Błąd okna etykiety GLS");
                MessageBox.Show(
                    $"Nie udało się otworzyć okna etykiety.\n\n{ex.Message}",
                    "Błąd etykiety",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                LabelButton.IsEnabled = true;
                _isBusy = false;
            }
        }

        private async System.Threading.Tasks.Task RefreshParcelNumbersAfterLabelAsync(int consignmentId)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                using var gls = new GlsService(_glsConfig);
                var shipment = await gls.GetShipmentAsync(consignmentId);
                if (!shipment.Success || shipment.Consignment == null)
                {
                    Warning(
                        shipment.ErrorMessage ?? "Nie udało się odczytać numerów listu po etykiecie GLS.",
                        "GlsWaybillDialog");
                    return;
                }

                var existingId = Consignment.ExistingId;
                Consignment = shipment.Consignment;
                Consignment.ExistingId = existingId ?? consignmentId;
                Consignment.NormalizeLimits();

                IssuedParcelNumbers = SubiektPrzesylka.NormalizeNrListu(Consignment.GetParcelNumbersDisplay());
                if (!string.IsNullOrWhiteSpace(IssuedParcelNumbers))
                {
                    _existingNrListu = IssuedParcelNumbers;
                }
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się odczytać numerów listu po etykiecie: {ex.Message}", "GlsWaybillDialog");
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            if (_hasIssuedLabel)
            {
                MessageBox.Show(
                    "Edycja przesyłki jest zablokowana, ponieważ dla tej przesyłki został już nadany numer listu.",
                    "Edycja zablokowana",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            RedistributeLongName1();
            if (!int.TryParse((ParcelCountTextBox.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parcelCount)
                || parcelCount < 1
                || parcelCount > GlsConsignment.MaxParcelCount)
            {
                MessageBox.Show(
                    $"Ilość paczek musi być liczbą od 1 do {GlsConsignment.MaxParcelCount}.",
                    "Ilość paczek",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                ParcelCountTextBox.Focus();
                return;
            }

            ApplyFormToConsignment();
            if (!Consignment.TryValidate(out var error))
            {
                MessageBox.Show(error, "Brak danych do przygotowalni", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var referenceText = string.IsNullOrWhiteSpace(Consignment.References)
                ? ""
                : $"\n\nDokument: {Consignment.References}";
            var confirm = MessageBox.Show(
                "Za chwilę usuniesz przesyłkę z przygotowalni GLS oraz lokalny wpis cache w Gryzaku."
                + referenceText
                + "\n\nTej operacji nie można cofnąć. Użyj tej opcji tylko wtedy, gdy na pewno wiesz, co robisz."
                + "\n\nCzy chcesz przejść dalej?",
                "Usuń z przygotowalni",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            var finalConfirm = MessageBox.Show(
                "To jest ostateczne potwierdzenie usunięcia przesyłki z przygotowalni GLS."
                + referenceText
                + "\n\nPotwierdź tylko wtedy, gdy świadomie chcesz usunąć ten dokument z przygotowalni.",
                "Potwierdź usunięcie",
                MessageBoxButton.YesNo,
                MessageBoxImage.Stop,
                MessageBoxResult.No);

            if (finalConfirm != MessageBoxResult.Yes)
            {
                return;
            }

            DeleteRequested = true;
            DialogResult = true;
            Close();
        }

        private void Name1TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            RedistributeLongName1();
        }

        private void RedistributeLongName1()
        {
            var name1 = (Name1TextBox.Text ?? "").Trim();
            if (name1.Length <= GlsConsignment.NameFieldMaxLength)
            {
                return;
            }

            GlsConsignment.SplitName(name1, out var part1, out var part2, out var part3);
            Name1TextBox.Text = part1;
            Name2TextBox.Text = part2;
            Name3TextBox.Text = part3;
        }

        private void ApplyFormToConsignment()
        {
            Consignment.Name1 = (Name1TextBox.Text ?? "").Trim();
            Consignment.Name2 = (Name2TextBox.Text ?? "").Trim();
            Consignment.Name3 = (Name3TextBox.Text ?? "").Trim();
            Consignment.Street = (StreetTextBox.Text ?? "").Trim();
            Consignment.Country = (CountryTextBox.Text ?? "").Trim();
            Consignment.ZipCode = (ZipCodeTextBox.Text ?? "").Trim();
            Consignment.City = (CityTextBox.Text ?? "").Trim();
            Consignment.Phone = (PhoneTextBox.Text ?? "").Trim();
            Consignment.Contact = (ContactTextBox.Text ?? "").Trim();
            Consignment.References = (ReferencesTextBox.Text ?? "").Trim();
            Consignment.Notes = (NotesTextBox.Text ?? "").Trim();
            Consignment.CashOnDelivery = CashOnDeliveryCheckBox.IsChecked == true;
            Consignment.ParcelWeight = string.IsNullOrWhiteSpace(WeightTextBox.Text)
                ? GlsConsignment.DefaultParcelWeightKg
                : WeightTextBox.Text.Trim();

            if (!int.TryParse((ParcelCountTextBox.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parcelCount)
                || parcelCount < 1)
            {
                parcelCount = 1;
            }

            Consignment.SetParcelCount(parcelCount);
            foreach (var parcel in Consignment.Parcels)
            {
                parcel.Reference = Consignment.References;
                parcel.Weight = Consignment.ParcelWeight;
            }
        }

        private static bool HasIssuedLabel(GlsConsignment consignment, string? existingNrListu)
        {
            var parcelNumbers = consignment?.GetParcelNumbersDisplay() ?? "";
            var combined = SubiektPrzesylka.NormalizeNrListu(
                GlsSubiektSync.CombineNrListu(existingNrListu, parcelNumbers));
            return !string.IsNullOrWhiteSpace(combined);
        }

        private void ApplyIssuedLabelLock()
        {
            if (!_hasIssuedLabel)
            {
                return;
            }

            Title = "Podgląd przesyłki w przygotowalni GLS";
            TitleText.Text = "Podgląd przesyłki w przygotowalni GLS";
            IssuedLabelNotice.Visibility = Visibility.Visible;
            FormPanel.IsEnabled = false;
            SubmitButton.Visibility = Visibility.Collapsed;
            SubmitButton.IsEnabled = false;
        }
    }
}
