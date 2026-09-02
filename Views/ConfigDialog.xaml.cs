using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Gryzak.Models;
using Gryzak.Services;

namespace Gryzak.Views
{
    public partial class ConfigDialog : Window
    {
        private readonly ConfigService _configService;
        private readonly ApiConfig _currentConfig;
        private bool _uiIsProduction;
        private bool _suppressEnvironmentChange;

        public ConfigDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadConfig();
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
                LoadUiFromCurrentEnvironment();
                UpdateEnvironmentLabels();
                UpdateUrlPreviews();
            }
            finally
            {
                _suppressEnvironmentChange = false;
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
            UpdateUrlPreviews();
            TestStatusText.Text = "";
        }

        private void LoadUiFromCurrentEnvironment()
        {
            var env = GetUiEnvironment();
            ApiUrlTextBox.Text = env.ApiUrl ?? "";
            ApiTokenPasswordBox.Password = env.ApiToken ?? "";
            ApiTimeoutTextBox.Text = env.ApiTimeout.ToString();
            OrderListEndpointTextBox.Text = env.OrderListEndpoint ?? "";
            OrderDetailsEndpointTextBox.Text = env.OrderDetailsEndpoint ?? "";
        }

        private void FlushUiToCurrentEnvironment()
        {
            var env = GetUiEnvironment();
            env.ApiUrl = ApiUrlTextBox.Text.Trim();
            env.ApiToken = ApiTokenPasswordBox.Password;
            env.ApiTimeout = int.TryParse(ApiTimeoutTextBox.Text, out var timeout) ? timeout : 30;
            env.OrderListEndpoint = OrderListEndpointTextBox.Text.Trim();
            env.OrderDetailsEndpoint = OrderDetailsEndpointTextBox.Text.Trim();
        }

        private ShopEnvironmentSettings GetUiEnvironment()
        {
            return _uiIsProduction ? _currentConfig.Production : _currentConfig.Test;
        }

        private void UpdateEnvironmentLabels()
        {
            if (EnvironmentConfigHeader == null)
            {
                return;
            }

            EnvironmentConfigHeader.Text = _uiIsProduction ? "Konfiguracja: Produkcja" : "Konfiguracja: Test";
        }

        private void UpdateUrlPreviews()
        {
            if (OrderListUrlPreview == null || ApiUrlTextBox == null)
            {
                return;
            }

            var baseUrl = ApiUrlTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                OrderListUrlPreview.Text = "Wprowadź główny URL API";
                OrderDetailsUrlPreview.Text = "Wprowadź główny URL API";
                return;
            }

            var baseUrlClean = baseUrl.TrimEnd('/');
            var listEndpoint = OrderListEndpointTextBox.Text.Trim();
            var detailsEndpoint = OrderDetailsEndpointTextBox.Text.Trim();

            var listUrl = baseUrlClean + listEndpoint;
            var separator = listEndpoint.Contains('?') ? "&" : "?";
            OrderListUrlPreview.Text = $"Pełny URL: {listUrl}{separator}page=1";

            var detailsUrl = baseUrlClean + detailsEndpoint.Replace("{order_id}", "123").Replace("{id}", "123");
            OrderDetailsUrlPreview.Text = $"Pełny URL: {detailsUrl}";
        }

        private void ApiUrl_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => UpdateUrlPreviews();

        private void ApiTimeout_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (int.TryParse(ApiTimeoutTextBox.Text, out var timeout))
            {
                ApiTimeoutTextBox.Background = timeout < 5 || timeout > 300
                    ? System.Windows.Media.Brushes.LightPink
                    : System.Windows.Media.Brushes.White;
            }
        }

        private void OrderListEndpoint_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => UpdateUrlPreviews();

        private void OrderDetailsEndpoint_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => UpdateUrlPreviews();

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            FlushUiToCurrentEnvironment();
            var env = GetUiEnvironment();

            TestConnectionButton.IsEnabled = false;
            TestStatusText.Text = "🔄 Testowanie połączenia...";
            TestStatusText.Foreground = System.Windows.Media.Brushes.Blue;

            try
            {
                if (string.IsNullOrWhiteSpace(env.ApiUrl))
                {
                    TestStatusText.Text = "❌ URL API jest wymagany";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(env.ApiTimeout);
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                if (!string.IsNullOrWhiteSpace(env.ApiToken))
                {
                    client.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", env.ApiToken);
                }

                var baseUrl = env.ApiUrl.TrimEnd('/');
                var endpoint = env.OrderListEndpoint;
                var separator = endpoint.Contains('?') ? "&" : "?";
                var url = $"{baseUrl}{endpoint}{separator}page=1";

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(env.ApiTimeout));
                var response = await client.GetAsync(url, cts.Token);

                if (response.IsSuccessStatusCode)
                {
                    TestStatusText.Text = "✅ Połączenie z API działa poprawnie!";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Green;
                }
                else
                {
                    TestStatusText.Text = $"❌ Błąd API: {(int)response.StatusCode} {response.ReasonPhrase}";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
                }
            }
            catch (TaskCanceledException)
            {
                TestStatusText.Text = "⏰ Timeout - połączenie przekroczyło limit czasu";
                TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
            }
            catch (HttpRequestException)
            {
                TestStatusText.Text = "🌐 Błąd sieci - sprawdź URL i połączenie internetowe";
                TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
            }
            catch (Exception ex)
            {
                TestStatusText.Text = $"❌ Błąd: {ex.Message}";
                TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
            }
            finally
            {
                TestConnectionButton.IsEnabled = true;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FlushUiToCurrentEnvironment();
                _currentConfig.Normalize();

                if (_currentConfig.Test.ApiTimeout < 5 || _currentConfig.Test.ApiTimeout > 300
                    || _currentConfig.Production.ApiTimeout < 5 || _currentConfig.Production.ApiTimeout > 300)
                {
                    MessageBox.Show(
                        "Timeout musi być między 5 a 300 sekundami (oba środowiska).",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                _configService.SaveConfig(_currentConfig);
                MessageBox.Show(
                    "Konfiguracja została zapisana pomyślnie.",
                    "Sukces",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Błąd zapisywania konfiguracji: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Zresetować aktualnie edytowany profil środowiska do wartości domyślnych?",
                "Potwierdzenie",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            var defaults = new ShopEnvironmentSettings();
            if (_uiIsProduction)
            {
                _currentConfig.Production = defaults;
            }
            else
            {
                defaults.ApiUrl = "https://mikran.pl";
                _currentConfig.Test = defaults;
            }

            LoadUiFromCurrentEnvironment();
            UpdateUrlPreviews();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
