using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class SubiektSettingsDialog : Window
    {
        private readonly ConfigService _configService;
        private SubiektConfig _currentConfig;
        private ObservableCollection<UserItem> _users = new ObservableCollection<UserItem>();

        private class UserItem
        {
            public string UserName { get; set; } = "";
            public string DisplayName { get; set; } = "";
            public int Id { get; set; }
        }

        public SubiektSettingsDialog(ConfigService configService)
        {
            InitializeComponent();
            _configService = configService;
            _currentConfig = _configService.LoadSubiektConfig();
            LoadConfig();
        }

        private SubiektConfig BuildConfigFromUi()
        {
            var config = new SubiektConfig
            {
                ApiBaseUrl = ApiBaseUrlTextBox.Text.Trim(),
                ApiKey = ApiKeyPasswordBox.Password,
                ServerAddress = ServerAddressTextBox.Text.Trim(),
                DatabaseName = DatabaseNameTextBox.Text.Trim(),
                ServerUsername = ServerUsernameTextBox.Text.Trim(),
                ServerPassword = ServerPasswordBox.Password,
                User = UserComboBox.Text.Trim(),
                Password = PasswordBox.Password,
                GtProdukt = _currentConfig.GtProdukt,
                AuthenticationMode = _currentConfig.AuthenticationMode,
                LaunchDopasujOperatora = _currentConfig.LaunchDopasujOperatora,
                LaunchTryb = _currentConfig.LaunchTryb,
                AutoReleaseLicenseTimeoutMinutes = _currentConfig.AutoReleaseLicenseTimeoutMinutes,
                DiscountCalculationMode = _currentConfig.DiscountCalculationMode,
                CalculateFromGrossPrices = _currentConfig.CalculateFromGrossPrices,
                DiscountRoundingMode = _currentConfig.DiscountRoundingMode
            };

            if (GtProduktComboBox.SelectedValue is string gtProduktStr && int.TryParse(gtProduktStr, out int gtProdukt))
                config.GtProdukt = gtProdukt;

            if (AuthenticationModeComboBox.SelectedValue is string authModeStr && int.TryParse(authModeStr, out int authMode))
                config.AuthenticationMode = authMode;

            if (LaunchDopasujComboBox.SelectedValue is string dopasujStr && int.TryParse(dopasujStr, out int dopasuj))
                config.LaunchDopasujOperatora = dopasuj;

            if (LaunchTrybComboBox.SelectedValue is string trybStr && int.TryParse(trybStr, out int tryb))
                config.LaunchTryb = tryb;

            var selectedDiscountMode = DiscountModeComboBox.SelectedValue as string;
            config.DiscountCalculationMode = string.IsNullOrWhiteSpace(selectedDiscountMode) ? "percent" : selectedDiscountMode;

            var selectedPriceMode = PriceCalculationModeComboBox.SelectedValue as string;
            config.CalculateFromGrossPrices = selectedPriceMode == "gross";

            var selectedRoundingMode = DiscountRoundingModeComboBox.SelectedValue as string;
            config.DiscountRoundingMode = string.IsNullOrWhiteSpace(selectedRoundingMode) ? "percent" : selectedRoundingMode;

            if (int.TryParse(AutoReleaseLicenseTimeoutTextBox.Text.Trim(), out int timeoutMinutes))
            {
                config.AutoReleaseLicenseTimeoutMinutes = timeoutMinutes < 0 ? 0 : timeoutMinutes;
            }

            return config;
        }

        private void LoadConfig()
        {
            ApiBaseUrlTextBox.Text = _currentConfig.ApiBaseUrl ?? "";
            ApiKeyPasswordBox.Password = _currentConfig.ApiKey ?? "";
            ServerAddressTextBox.Text = _currentConfig.ServerAddress ?? "";
            DatabaseNameTextBox.Text = _currentConfig.DatabaseName ?? "";
            ServerUsernameTextBox.Text = _currentConfig.ServerUsername ?? "";
            ServerPasswordBox.Password = _currentConfig.ServerPassword ?? "";
            
            UserComboBox.ItemsSource = _users;
            
            string savedUser = _currentConfig.User ?? "";
            if (!string.IsNullOrEmpty(savedUser))
            {
                UserComboBox.Text = savedUser;
            }
            
            PasswordBox.Password = _currentConfig.Password;
            GtProduktComboBox.SelectedValue = _currentConfig.GtProdukt.ToString();
            AuthenticationModeComboBox.SelectedValue = _currentConfig.AuthenticationMode.ToString();
            LaunchDopasujComboBox.SelectedValue = _currentConfig.LaunchDopasujOperatora.ToString();
            LaunchTrybComboBox.SelectedValue = _currentConfig.LaunchTryb.ToString();
            AutoReleaseLicenseTimeoutTextBox.Text = _currentConfig.AutoReleaseLicenseTimeoutMinutes.ToString();

            if (string.IsNullOrWhiteSpace(_currentConfig.DiscountCalculationMode))
            {
                _currentConfig.DiscountCalculationMode = "percent";
            }
            DiscountModeComboBox.SelectedValue = _currentConfig.DiscountCalculationMode;

            PriceCalculationModeComboBox.SelectedValue = _currentConfig.CalculateFromGrossPrices ? "gross" : "net";

            if (string.IsNullOrWhiteSpace(_currentConfig.DiscountRoundingMode))
            {
                _currentConfig.DiscountRoundingMode = "percent";
            }
            DiscountRoundingModeComboBox.SelectedValue = _currentConfig.DiscountRoundingMode;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            _currentConfig = BuildConfigFromUi();

            if (_currentConfig.AutoReleaseLicenseTimeoutMinutes < 0)
            {
                MessageBox.Show("Czas nieaktywności nie może być ujemny. Ustawiono wartość 0 (wyłączone).", "Ostrzeżenie", MessageBoxButton.OK, MessageBoxImage.Warning);
                _currentConfig.AutoReleaseLicenseTimeoutMinutes = 0;
            }

            if (!int.TryParse(AutoReleaseLicenseTimeoutTextBox.Text.Trim(), out _))
            {
                MessageBox.Show("Nieprawidłowa wartość czasu nieaktywności. Ustawiono wartość 0 (wyłączone).", "Ostrzeżenie", MessageBoxButton.OK, MessageBoxImage.Warning);
                _currentConfig.AutoReleaseLicenseTimeoutMinutes = 0;
            }

            try
            {
                _configService.SaveSubiektConfig(_currentConfig);
                MessageBox.Show("Ustawienia Subiekt GT zostały zapisane pomyślnie.", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Nie udało się zapisać ustawień:\n\n{ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !char.IsDigit(e.Text, e.Text.Length - 1);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var config = BuildConfigFromUi();

                if (string.IsNullOrWhiteSpace(config.ApiBaseUrl))
                {
                    MessageBox.Show("Proszę podać URL API Subiekt (z /api/v1).", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                TestConnectionButton.IsEnabled = false;
                TestConnectionButton.Content = "⏳ Testowanie...";
                Mouse.OverrideCursor = Cursors.Wait;

                var api = new SubiektApiService(_configService);
                var (ok, message) = await api.TestConnectionAsync(config);

                if (ok)
                {
                    Info($"Połączenie z API Subiekt udane: {message}", "SubiektSettings");
                    MessageBox.Show(
                        $"Połączenie z API Subiekt zakończone pomyślnie!\n\n{message}",
                        "Sukces",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektSettings", "Błąd połączenia z API Subiekt");
                MessageBox.Show(
                    $"Nie udało się połączyć z API Subiekt.\n\nBłąd: {ex.Message}",
                    "Błąd połączenia",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                TestConnectionButton.IsEnabled = true;
                TestConnectionButton.Content = "🔌 Testuj połączenie API";
                Mouse.OverrideCursor = null;
            }
        }

        private async void LoadUsersButton_Click(object sender, RoutedEventArgs e)
        {
            var config = BuildConfigFromUi();

            if (string.IsNullOrWhiteSpace(config.ApiBaseUrl))
            {
                MessageBox.Show("Proszę podać URL API Subiekt (z /api/v1).", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            LoadUsersButton.IsEnabled = false;
            LoadUsersButton.Content = "⏳ Pobieranie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                var api = new SubiektApiService(_configService);
                var usersList = await api.GetUsersAsync(config);

                _users.Clear();
                foreach (var user in usersList)
                {
                    string displayName = $"{user.Uz_Nazwisko} {user.Uz_Imie}".Trim();
                    _users.Add(new UserItem
                    {
                        Id = user.Uz_Id,
                        UserName = displayName,
                        DisplayName = displayName
                    });
                }

                string savedUser = _currentConfig.User ?? "";
                if (!string.IsNullOrEmpty(savedUser))
                {
                    UserComboBox.Text = savedUser;
                }

                Info($"Pobrano {usersList.Count} użytkowników z API Subiekt.", "SubiektSettings");
                MessageBox.Show(
                    $"Lista użytkowników została pobrana ({usersList.Count} użytkowników).\n\nWybierz użytkownika z listy powyżej.",
                    "Sukces",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Nie udało się pobrać listy użytkowników.\n\nBłąd: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                
                Error(ex, "SubiektSettings");
            }
            finally
            {
                LoadUsersButton.IsEnabled = true;
                LoadUsersButton.Content = "📋 Pobierz listę użytkowników (API)";
                Mouse.OverrideCursor = null;
            }
        }

        private async void TestSubiektButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _currentConfig = BuildConfigFromUi();

                if (string.IsNullOrWhiteSpace(_currentConfig.ServerAddress))
                {
                    MessageBox.Show("Proszę podać adres serwera (Sfera).", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(_currentConfig.DatabaseName))
                {
                    MessageBox.Show("Proszę podać nazwę bazy danych.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                try
                {
                    _configService.SaveSubiektConfig(_currentConfig);
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się zapisać ustawień przed testem: {ex.Message}", "SubiektSettings");
                }

                TestSubiektButton.IsEnabled = false;
                TestSubiektButton.Content = "⏳ Uruchamianie...";
                Mouse.OverrideCursor = Cursors.Wait;

                await Task.Delay(100);
                
                _ = Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        var subiektService = new SubiektService();
                        subiektService.TestujUruchomienieSubiekta();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Błąd podczas testowego uruchomienia Subiekta GT:\n\n{ex.Message}",
                            "Błąd",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        Error(ex, "SubiektSettings", "Błąd podczas testowego uruchomienia");
                    }
                    finally
                    {
                        TestSubiektButton.IsEnabled = true;
                        TestSubiektButton.Content = "🚀 Testuj uruchomienie Subiekta GT";
                        Mouse.OverrideCursor = null;
                    }
                }), System.Windows.Threading.DispatcherPriority.Normal);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Błąd:\n\n{ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Error(ex, "SubiektSettings");
                
                TestSubiektButton.IsEnabled = true;
                TestSubiektButton.Content = "🚀 Testuj uruchomienie Subiekta GT";
                Mouse.OverrideCursor = null;
            }
        }
    }
}
