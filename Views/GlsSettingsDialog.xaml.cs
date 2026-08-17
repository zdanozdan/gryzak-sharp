using System;
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
        private GlsConfig _currentConfig;

        public GlsSettingsDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadGlsConfig();
            LoadConfig();
        }

        private void LoadConfig()
        {
            UserNameTextBox.Text = _currentConfig.UserName ?? "";
            PasswordBox.Password = _currentConfig.Password ?? "";
            TestApiUrlTextBox.Text = string.IsNullOrWhiteSpace(_currentConfig.TestApiUrl)
                ? GlsConfig.DefaultTestApiUrl
                : _currentConfig.TestApiUrl;
            ProductionApiUrlTextBox.Text = string.IsNullOrWhiteSpace(_currentConfig.ProductionApiUrl)
                ? GlsConfig.DefaultProductionApiUrl
                : _currentConfig.ProductionApiUrl;
            TimeoutTextBox.Text = _currentConfig.TimeoutSeconds.ToString();

            TestEnvironmentRadio.IsChecked = !_currentConfig.UseProduction;
            ProductionEnvironmentRadio.IsChecked = _currentConfig.UseProduction;
            UpdateActiveUrlPreview();
        }

        private void Environment_Changed(object sender, RoutedEventArgs e)
        {
            UpdateActiveUrlPreview();
        }

        private void ApiUrl_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateActiveUrlPreview();
        }

        private void UpdateActiveUrlPreview()
        {
            if (ActiveUrlPreview == null)
            {
                return;
            }

            var config = GetConfigFromUI();
            ActiveUrlPreview.Text = config.GetActiveApiUrl();
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            var config = GetConfigFromUI();

            if (string.IsNullOrWhiteSpace(config.UserName) || string.IsNullOrWhiteSpace(config.Password))
            {
                SetTestStatus("Login i hasło są wymagane.", Brushes.Red);
                MessageBox.Show("Proszę podać login i hasło GLS.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
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

            if (config.TimeoutSeconds < 5 || config.TimeoutSeconds > 300)
            {
                MessageBox.Show("Timeout musi być między 5 a 300 sekundami.", "Błąd walidacji", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(config.TestApiUrl))
            {
                config.TestApiUrl = GlsConfig.DefaultTestApiUrl;
            }

            if (string.IsNullOrWhiteSpace(config.ProductionApiUrl))
            {
                config.ProductionApiUrl = GlsConfig.DefaultProductionApiUrl;
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
            return new GlsConfig
            {
                UserName = UserNameTextBox.Text.Trim(),
                Password = PasswordBox.Password,
                UseProduction = ProductionEnvironmentRadio.IsChecked == true,
                TestApiUrl = TestApiUrlTextBox.Text.Trim(),
                ProductionApiUrl = ProductionApiUrlTextBox.Text.Trim(),
                TimeoutSeconds = int.TryParse(TimeoutTextBox.Text.Trim(), out var timeout) ? timeout : 30
            };
        }

        private void SetTestStatus(string message, Brush color)
        {
            TestStatusText.Text = message;
            TestStatusText.Foreground = color;
        }
    }
}
