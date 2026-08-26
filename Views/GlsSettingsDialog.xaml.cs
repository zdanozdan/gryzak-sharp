using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class GlsSettingsDialog : Window
    {
        private readonly ConfigService _configService;
        private readonly GlsConfig _currentConfig;
        private bool _uiIsProduction;
        private bool _suppressEnvironmentChange;

        public GlsSettingsDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadGlsConfig();
            LoadConfig();
        }

        private void LoadConfig()
        {
            _suppressEnvironmentChange = true;
            try
            {
                _uiIsProduction = _currentConfig.UseProduction;
                TestEnvironmentRadio.IsChecked = !_uiIsProduction;
                ProductionEnvironmentRadio.IsChecked = _uiIsProduction;
                GlsPanelEnabledCheckBox.IsChecked = _currentConfig.GlsPanelEnabled;
                LoadUiFromCurrentEnvironment();
                UpdateEnvironmentLabels();
                UpdateActiveUrlPreview();
                LoadPrinters();
            }
            finally
            {
                _suppressEnvironmentChange = false;
            }
        }

        private void LoadPrinters()
        {
            var printers = RawPrinterHelper.GetInstalledPrinterNames().ToList();
            if (!string.IsNullOrWhiteSpace(_currentConfig.LabelPrinterName)
                && !printers.Contains(_currentConfig.LabelPrinterName, StringComparer.OrdinalIgnoreCase))
            {
                printers.Insert(0, _currentConfig.LabelPrinterName);
            }

            LabelPrinterComboBox.ItemsSource = printers;
            var preferred = RawPrinterHelper.PickPreferredPrinter(_currentConfig.LabelPrinterName, printers);
            if (preferred != null)
            {
                LabelPrinterComboBox.SelectedItem = preferred;
            }
        }

        private void Environment_Changed(object sender, RoutedEventArgs e)
        {
            if (_suppressEnvironmentChange || ApiUrlTextBox == null)
            {
                return;
            }

            FlushUiToCurrentEnvironment();
            _uiIsProduction = ProductionEnvironmentRadio.IsChecked == true;
            LoadUiFromCurrentEnvironment();
            UpdateEnvironmentLabels();
            UpdateActiveUrlPreview();
            SetTestStatus("", Brushes.Gray);
        }

        private void ApiUrl_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateActiveUrlPreview();
        }

        private void UpdateEnvironmentLabels()
        {
            if (EnvironmentConfigHeader == null || ApiUrlLabel == null)
            {
                return;
            }

            if (_uiIsProduction)
            {
                EnvironmentConfigHeader.Text = "Konfiguracja: Produkcja";
                ApiUrlLabel.Text = "URL API (produkcja):";
            }
            else
            {
                EnvironmentConfigHeader.Text = "Konfiguracja: Test";
                ApiUrlLabel.Text = "URL API (test):";
            }
        }

        private void UpdateActiveUrlPreview()
        {
            if (ActiveUrlPreview == null || ApiUrlTextBox == null)
            {
                return;
            }

            var url = ApiUrlTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(url))
            {
                url = ProductionEnvironmentRadio.IsChecked == true
                    ? GlsConfig.DefaultProductionApiUrl
                    : GlsConfig.DefaultTestApiUrl;
            }

            ActiveUrlPreview.Text = url;
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            var config = GetConfigFromUI();

            if (string.IsNullOrWhiteSpace(config.UserName) || string.IsNullOrWhiteSpace(config.Password))
            {
                SetTestStatus("Login i hasło są wymagane.", Brushes.Red);
                MessageBox.Show("Proszę podać login i hasło GLS dla wybranego środowiska.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(config.GetActiveApiUrl()))
            {
                SetTestStatus("URL API jest wymagany.", Brushes.Red);
                MessageBox.Show("Proszę podać URL API dla wybranego środowiska.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TestConnectionButton.IsEnabled = false;
            TestConnectionButton.Content = "⏳ Testowanie...";
            Mouse.OverrideCursor = Cursors.Wait;
            SetTestStatus($"Logowanie do środowiska {config.GetEnvironmentName()}...", Brushes.DodgerBlue);

            try
            {
                using var glsService = new GlsService(config);
                var result = await glsService.TestConnectionAsync();

                if (result.Success)
                {
                    var message = $"Logowanie do GLS ({result.EnvironmentName}) zakończone pomyślnie.";
                    SetTestStatus($"✅ {message}", Brushes.Green);
                    Info(message, "GlsSettings");
                    MessageBox.Show(message, "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    var error = result.ErrorMessage ?? "Nieznany błąd logowania GLS.";
                    SetTestStatus($"❌ {error}", Brushes.Red);
                    Warning(error, "GlsSettings");
                    MessageBox.Show(
                        $"Nie udało się zalogować do GLS ({result.EnvironmentName}).\n\n{error}",
                        "Błąd logowania",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                SetTestStatus($"❌ {ex.Message}", Brushes.Red);
                Error(ex, "GlsSettings", "Błąd testu logowania GLS");
                MessageBox.Show($"Błąd testu logowania GLS:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                TestConnectionButton.IsEnabled = true;
                TestConnectionButton.Content = "🧪 Testuj logowanie";
                Mouse.OverrideCursor = null;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var config = GetConfigFromUI();

            if (config.Test.TimeoutSeconds < 5 || config.Test.TimeoutSeconds > 300
                || config.Production.TimeoutSeconds < 5 || config.Production.TimeoutSeconds > 300)
            {
                MessageBox.Show("Timeout musi być między 5 a 300 sekundami.", "Błąd walidacji", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(config.Test.ApiUrl))
            {
                config.Test.ApiUrl = GlsConfig.DefaultTestApiUrl;
            }

            if (string.IsNullOrWhiteSpace(config.Production.ApiUrl))
            {
                config.Production.ApiUrl = GlsConfig.DefaultProductionApiUrl;
            }

            try
            {
                _configService.SaveGlsConfig(config);
                MessageBox.Show("Ustawienia GLS zostały zapisane pomyślnie.", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Nie udało się zapisać ustawień GLS:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private GlsConfig GetConfigFromUI()
        {
            FlushUiToCurrentEnvironment();
            _currentConfig.UseProduction = ProductionEnvironmentRadio.IsChecked == true;
            _currentConfig.GlsPanelEnabled = GlsPanelEnabledCheckBox.IsChecked == true;
            _currentConfig.LabelPrinterName = (LabelPrinterComboBox.SelectedItem as string)?.Trim()
                ?? LabelPrinterComboBox.Text?.Trim()
                ?? "";
            return _currentConfig;
        }

        private void FlushUiToCurrentEnvironment()
        {
            var env = GetUiEnvironment();
            env.ApiUrl = ApiUrlTextBox.Text.Trim();
            env.UserName = UserNameTextBox.Text.Trim();
            env.Password = PasswordBox.Password;
            env.TimeoutSeconds = int.TryParse(TimeoutTextBox.Text.Trim(), out var timeout) ? timeout : 30;
        }

        private void LoadUiFromCurrentEnvironment()
        {
            var env = GetUiEnvironment();
            var defaultUrl = _uiIsProduction ? GlsConfig.DefaultProductionApiUrl : GlsConfig.DefaultTestApiUrl;
            ApiUrlTextBox.Text = string.IsNullOrWhiteSpace(env.ApiUrl) ? defaultUrl : env.ApiUrl;
            UserNameTextBox.Text = env.UserName ?? "";
            PasswordBox.Password = env.Password ?? "";
            TimeoutTextBox.Text = env.TimeoutSeconds.ToString();
        }

        private GlsEnvironmentSettings GetUiEnvironment()
        {
            return _uiIsProduction ? _currentConfig.Production : _currentConfig.Test;
        }

        private void SetTestStatus(string message, Brush color)
        {
            TestStatusText.Text = message;
            TestStatusText.Foreground = color;
        }
    }
}
