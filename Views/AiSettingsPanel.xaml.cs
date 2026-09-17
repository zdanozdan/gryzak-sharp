using System;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Gryzak.Models;
using Gryzak.Services;
using Gryzak.Services.Llm;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class AiSettingsPanel : UserControl
    {
        private static readonly string[] GeminiModels =
        {
            "gemini-3.5-flash-lite",
            "gemini-3.5-flash",
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-2.0-flash"
        };

        private readonly ConfigService _configService;
        private AiConfig _currentConfig;
        private bool _suppressProviderChange;
        private CancellationTokenSource? _testCts;

        public AiSettingsPanel(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadAiConfig();
            InitializeProviders();
            LoadConfig();
        }

        public bool TrySave()
        {
            try
            {
                var config = GetConfigFromUI();
                config.Normalize();

                if (config.TimeoutSeconds < 5 || config.TimeoutSeconds > 300)
                {
                    MessageBox.Show(
                        "Timeout AI musi być między 5 a 300 sekundami.",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return false;
                }

                if (config.Temperature < 0 || config.Temperature > 2)
                {
                    MessageBox.Show(
                        "Temperatura musi być w zakresie 0–2.",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(config.Model))
                {
                    MessageBox.Show(
                        "Nazwa modelu jest wymagana.",
                        "Błąd walidacji",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return false;
                }

                _configService.SaveAiConfig(config);
                _currentConfig = config;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Błąd zapisywania ustawień AI: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        private void InitializeProviders()
        {
            _suppressProviderChange = true;
            ProviderComboBox.Items.Clear();
            ProviderComboBox.Items.Add(new ComboBoxItem
            {
                Content = AiProviders.GetDisplayName(AiProviders.GoogleGemini),
                Tag = AiProviders.GoogleGemini
            });
            ProviderComboBox.SelectedIndex = 0;
            _suppressProviderChange = false;
        }

        private void LoadConfig()
        {
            SelectProvider(_currentConfig.Provider);
            RefreshModelSuggestions();
            ModelComboBox.Text = _currentConfig.Model ?? AiConfig.DefaultGeminiModel;
            ApiKeyPasswordBox.Password = _currentConfig.ApiKey ?? "";
            ApiBaseUrlTextBox.Text = _currentConfig.ApiBaseUrl ?? "";
            TemperatureTextBox.Text = _currentConfig.Temperature.ToString("0.##", CultureInfo.InvariantCulture);
            MaxOutputTokensTextBox.Text = _currentConfig.MaxOutputTokens.ToString(CultureInfo.InvariantCulture);
            TimeoutTextBox.Text = _currentConfig.TimeoutSeconds.ToString(CultureInfo.InvariantCulture);
            UpdateApiBaseUrlPreview();
        }

        private void SelectProvider(string? provider)
        {
            _suppressProviderChange = true;
            var match = AiProviders.GoogleGemini;
            if (!string.IsNullOrWhiteSpace(provider))
            {
                match = provider.Trim().ToLowerInvariant();
            }

            foreach (ComboBoxItem item in ProviderComboBox.Items)
            {
                if (string.Equals(item.Tag as string, match, StringComparison.OrdinalIgnoreCase))
                {
                    ProviderComboBox.SelectedItem = item;
                    _suppressProviderChange = false;
                    return;
                }
            }

            ProviderComboBox.SelectedIndex = 0;
            _suppressProviderChange = false;
        }

        private string GetSelectedProvider()
        {
            if (ProviderComboBox.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                return tag;
            }

            return AiProviders.GoogleGemini;
        }

        private void RefreshModelSuggestions()
        {
            var current = ModelComboBox.Text;
            ModelComboBox.Items.Clear();
            foreach (var model in GeminiModels)
            {
                ModelComboBox.Items.Add(model);
            }

            if (!string.IsNullOrWhiteSpace(current))
            {
                ModelComboBox.Text = current;
            }
        }

        private AiConfig GetConfigFromUI()
        {
            return new AiConfig
            {
                Provider = GetSelectedProvider(),
                Model = (ModelComboBox.Text ?? "").Trim(),
                ApiKey = ApiKeyPasswordBox.Password,
                ApiBaseUrl = (ApiBaseUrlTextBox.Text ?? "").Trim(),
                Temperature = double.TryParse(
                    TemperatureTextBox.Text?.Trim(),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var temperature)
                    ? temperature
                    : 0.7,
                MaxOutputTokens = int.TryParse(
                    MaxOutputTokensTextBox.Text?.Trim(),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var maxTokens)
                    ? maxTokens
                    : 8192,
                TimeoutSeconds = int.TryParse(
                    TimeoutTextBox.Text?.Trim(),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var timeout)
                    ? timeout
                    : 60
            };
        }

        private void ProviderComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressProviderChange || !IsLoaded)
            {
                return;
            }

            RefreshModelSuggestions();
            UpdateApiBaseUrlPreview();
        }

        private void ApiBaseUrl_TextChanged(object sender, TextChangedEventArgs e) => UpdateApiBaseUrlPreview();

        private void UpdateApiBaseUrlPreview()
        {
            if (ApiBaseUrlPreview == null)
            {
                return;
            }

            var draft = GetConfigFromUI();
            draft.Normalize();
            var url = draft.GetEffectiveApiBaseUrl();
            ApiBaseUrlPreview.Text = string.IsNullOrWhiteSpace(url)
                ? "Brak domyślnego URL dla wybranego dostawcy — uzupełnij pole powyżej."
                : $"Używany URL: {url}";
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            var config = GetConfigFromUI();
            config.Normalize();

            if (string.IsNullOrWhiteSpace(config.ApiKey))
            {
                TestStatusText.Text = "❌ Podaj klucz API przed testem.";
                TestStatusText.Foreground = Brushes.Red;
                return;
            }

            if (string.IsNullOrWhiteSpace(config.Model))
            {
                TestStatusText.Text = "❌ Podaj nazwę modelu przed testem.";
                TestStatusText.Foreground = Brushes.Red;
                return;
            }

            _testCts?.Cancel();
            _testCts?.Dispose();
            _testCts = new CancellationTokenSource();

            TestConnectionButton.IsEnabled = false;
            TestStatusText.Text = $"🔄 Testowanie {AiProviders.GetDisplayName(config.Provider)} / {config.Model}…";
            TestStatusText.Foreground = Brushes.Blue;

            try
            {
                using var client = LlmClientFactory.Create(config);
                var result = await client.CompleteAsync(
                    new LlmCompletionRequest
                    {
                        SystemPrompt = "Odpowiadasz krótko i po polsku.",
                        Messages =
                        {
                            LlmMessage.User(
                                "To jest test połączenia z API. Odpowiedz dokładnie jednym słowem: OK")
                        },
                        Temperature = 0,
                        MaxOutputTokens = 32
                    },
                    _testCts.Token);

                if (result.Success)
                {
                    var preview = result.Text.Length > 120
                        ? result.Text.Substring(0, 120) + "…"
                        : result.Text;
                    TestStatusText.Text =
                        $"✅ Połączenie OK ({result.Duration.TotalSeconds:0.0}s)\n" +
                        $"Model: {result.Model}\n" +
                        $"Odpowiedź: {preview}";
                    TestStatusText.Foreground = Brushes.Green;
                    Info(
                        $"Test AI OK: {result.ProviderId}/{result.Model} — {result.Text}",
                        "AiSettings");
                }
                else
                {
                    TestStatusText.Text = $"❌ {result.ErrorMessage}";
                    TestStatusText.Foreground = Brushes.Red;
                    Warning($"Test AI nieudany: {result.ErrorMessage}", "AiSettings");
                }
            }
            catch (NotSupportedException ex)
            {
                TestStatusText.Text = $"❌ {ex.Message}";
                TestStatusText.Foreground = Brushes.Red;
            }
            catch (Exception ex)
            {
                TestStatusText.Text = $"❌ Błąd: {ex.Message}";
                TestStatusText.Foreground = Brushes.Red;
                Error(ex, "AiSettings", "Błąd testu AI");
            }
            finally
            {
                TestConnectionButton.IsEnabled = true;
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Zresetować ustawienia AI do wartości domyślnych?",
                "Potwierdzenie",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            _currentConfig = new AiConfig();
            _currentConfig.Normalize();
            LoadConfig();
            if (TestStatusText != null)
            {
                TestStatusText.Text = "";
            }
        }
    }
}
