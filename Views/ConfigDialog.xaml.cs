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
        private ApiConfig _currentConfig;

        public ConfigDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadConfig();
            LoadConfig();
            UpdateUrlPreviews();
        }

        private void LoadConfig()
        {
            ApiUrlTextBox.Text = _currentConfig.ApiUrl;
            ApiTokenPasswordBox.Password = _currentConfig.ApiToken;
            ApiTimeoutTextBox.Text = _currentConfig.ApiTimeout.ToString();
            OrderListEndpointTextBox.Text = _currentConfig.OrderListEndpoint;
            OrderDetailsEndpointTextBox.Text = _currentConfig.OrderDetailsEndpoint;
        }

        private void UpdateUrlPreviews()
        {
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

            var detailsUrl = baseUrlClean + detailsEndpoint.Replace("{id}", "123");
            OrderDetailsUrlPreview.Text = $"Pełny URL: {detailsUrl}";
        }

        private void ApiUrl_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateUrlPreviews();
        }

        private void ApiToken_PasswordChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            // PasswordChanged tylko do aktualizacji podglądu jeśli potrzebne
        }

        private void ApiTimeout_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Walidacja
            if (int.TryParse(ApiTimeoutTextBox.Text, out var timeout))
            {
                if (timeout < 5 || timeout > 300)
                {
                    ApiTimeoutTextBox.Background = System.Windows.Media.Brushes.LightPink;
                }
                else
                {
                    ApiTimeoutTextBox.Background = System.Windows.Media.Brushes.White;
                }
            }
        }

        private void OrderListEndpoint_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateUrlPreviews();
        }

        private void OrderDetailsEndpoint_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateUrlPreviews();
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            TestConnectionButton.IsEnabled = false;
            TestStatusText.Text = "🔄 Testowanie połączenia...";
            TestStatusText.Foreground = System.Windows.Media.Brushes.Blue;

            try
            {
                var config = GetConfigFromUI();
                
                if (string.IsNullOrWhiteSpace(config.ApiUrl))
                {
                    TestStatusText.Text = "❌ URL API jest wymagany";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
                    TestConnectionButton.IsEnabled = true;
                    return;
                }

                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(config.ApiTimeout);
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                
                if (!string.IsNullOrWhiteSpace(config.ApiToken))
                {
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", config.ApiToken);
                }

                var baseUrl = config.ApiUrl.TrimEnd('/');
                var endpoint = config.OrderListEndpoint;
                var separator = endpoint.Contains('?') ? "&" : "?";
                var url = $"{baseUrl}{endpoint}{separator}page=1";

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(config.ApiTimeout));
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
                var config = GetConfigFromUI();

                // Walidacja
                if (string.IsNullOrWhiteSpace(config.ApiUrl))
                {
                    MessageBox.Show(
                        "URL API jest wymagany do zapisania konfiguracji.",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (config.ApiTimeout < 5 || config.ApiTimeout > 300)
                {
                    MessageBox.Show(
                        "Timeout musi być między 5 a 300 sekundami.",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(config.OrderListEndpoint))
                {
                    config.OrderListEndpoint = "/orders";
                }

                if (string.IsNullOrWhiteSpace(config.OrderDetailsEndpoint))
                {
                    config.OrderDetailsEndpoint = "/orders";
                }

                _configService.SaveConfig(config);
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
                "Czy na pewno chcesz zresetować konfigurację do wartości domyślnych?",
                "Potwierdzenie",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _currentConfig = new ApiConfig();
                LoadConfig();
                UpdateUrlPreviews();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private ApiConfig GetConfigFromUI()
        {
            return new ApiConfig
            {
                ApiUrl = ApiUrlTextBox.Text.Trim(),
                ApiToken = ApiTokenPasswordBox.Password,
                ApiTimeout = int.TryParse(ApiTimeoutTextBox.Text, out var timeout) ? timeout : 30,
                OrderListEndpoint = OrderListEndpointTextBox.Text.Trim(),
                OrderDetailsEndpoint = OrderDetailsEndpointTextBox.Text.Trim()
            };
        }
    }
}

