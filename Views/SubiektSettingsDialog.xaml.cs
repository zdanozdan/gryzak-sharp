using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Gryzak.Models;
using Gryzak.Services;
using Microsoft.Data.SqlClient;

namespace Gryzak.Views
{
    public partial class SubiektSettingsDialog : Window
    {
        private readonly ConfigService _configService;
        private SubiektConfig _currentConfig;
        private ObservableCollection<UserItem> _users = new ObservableCollection<UserItem>();

        // Klasa pomocnicza do reprezentacji użytkownika
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

        private void LoadConfig()
        {
            ServerAddressTextBox.Text = _currentConfig.ServerAddress ?? "";
            ServerUsernameTextBox.Text = _currentConfig.ServerUsername ?? "";
            ServerPasswordBox.Password = _currentConfig.ServerPassword ?? "";
            
            // Ustaw źródło danych dla ComboBox
            UserComboBox.ItemsSource = _users;
            
            // Ustaw wybranego użytkownika jeśli istnieje w konfiguracji
            string savedUser = _currentConfig.User ?? "";
            if (!string.IsNullOrEmpty(savedUser))
            {
                UserComboBox.Text = savedUser;
            }
            
            PasswordBox.Password = _currentConfig.Password;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            _currentConfig.ServerAddress = ServerAddressTextBox.Text.Trim();
            _currentConfig.ServerUsername = ServerUsernameTextBox.Text.Trim();
            _currentConfig.ServerPassword = ServerPasswordBox.Password;
            _currentConfig.User = UserComboBox.Text.Trim();
            _currentConfig.Password = PasswordBox.Password;

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

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private async void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            string serverAddress = ServerAddressTextBox.Text.Trim();
            string username = ServerUsernameTextBox.Text.Trim();
            string password = ServerPasswordBox.Password;

            // Walidacja podstawowa
            if (string.IsNullOrWhiteSpace(serverAddress))
            {
                MessageBox.Show("Proszę podać adres serwera MSSQL.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Wyłącz przycisk podczas testowania
            TestConnectionButton.IsEnabled = false;
            TestConnectionButton.Content = "⏳ Testowanie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                // Utwórz connection string
                var builder = new SqlConnectionStringBuilder
                {
                    DataSource = serverAddress,
                    UserID = username,
                    Password = password,
                    ConnectTimeout = 10, // 10 sekund timeout
                    Encrypt = false // Dla starszych serwerów MSSQL
                };

                // Jeśli nie podano username/password, użyj Windows Authentication
                if (string.IsNullOrWhiteSpace(username))
                {
                    builder.IntegratedSecurity = true;
                }

                string connectionString = builder.ConnectionString;

                // Test połączenia asynchronicznie
                bool success = await Task.Run(() =>
                {
                    try
                    {
                        using (var connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            // Wykonaj prosty query aby sprawdzić czy połączenie działa
                            using (var command = new SqlCommand("SELECT @@VERSION", connection))
                            {
                                var version = command.ExecuteScalar();
                                Console.WriteLine($"[SubiektSettings] Połączenie z MSSQL udane. Wersja serwera: {version}");
                            }
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SubiektSettings] Błąd połączenia z MSSQL: {ex.Message}");
                        throw;
                    }
                });

                if (success)
                {
                    MessageBox.Show(
                        "Połączenie z serwerem MSSQL zakończone pomyślnie!",
                        "Sukces",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (SqlException sqlEx)
            {
                string errorMessage = "Nie udało się połączyć z serwerem MSSQL.\n\n";
                errorMessage += $"Błąd: {sqlEx.Message}";
                
                if (sqlEx.Number == 18456)
                {
                    errorMessage += "\n\nSprawdź poprawność nazwy użytkownika i hasła.";
                }
                else if (sqlEx.Number == -1 || sqlEx.Number == 2)
                {
                    errorMessage += "\n\nNie można nawiązać połączenia. Sprawdź:\n- Czy adres serwera jest poprawny\n- Czy serwer jest dostępny w sieci\n- Czy firewall nie blokuje połączenia";
                }

                MessageBox.Show(errorMessage, "Błąd połączenia", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Nie udało się połączyć z serwerem MSSQL.\n\nBłąd: {ex.Message}",
                    "Błąd połączenia",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                // Przywróć przycisk i kursor
                TestConnectionButton.IsEnabled = true;
                TestConnectionButton.Content = "🔌 Testuj połączenie";
                Mouse.OverrideCursor = null;
            }
        }

        private async void LoadUsersButton_Click(object sender, RoutedEventArgs e)
        {
            string serverAddress = ServerAddressTextBox.Text.Trim();
            string username = ServerUsernameTextBox.Text.Trim();
            string password = ServerPasswordBox.Password;

            // Walidacja podstawowa
            if (string.IsNullOrWhiteSpace(serverAddress))
            {
                MessageBox.Show("Proszę podać adres serwera MSSQL.", "Brak danych", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Wyłącz przycisk podczas pobierania
            LoadUsersButton.IsEnabled = false;
            LoadUsersButton.Content = "⏳ Pobieranie...";
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                // Utwórz connection string
                var builder = new SqlConnectionStringBuilder
                {
                    DataSource = serverAddress,
                    UserID = username,
                    Password = password,
                    ConnectTimeout = 10,
                    Encrypt = false
                };

                // Jeśli nie podano username/password, użyj Windows Authentication
                if (string.IsNullOrWhiteSpace(username))
                {
                    builder.IntegratedSecurity = true;
                }

                string connectionString = builder.ConnectionString;

                // Wykonaj zapytanie SQL asynchronicznie
                var usersList = await Task.Run(() =>
                {
                    var users = new System.Collections.Generic.List<UserItem>();
                    try
                    {
                        using (var connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            
                            string sqlQuery = @"
SELECT [uz_Id]
      ,[uz_Nazwisko]
      ,[uz_Imie]
      ,[uz_Status]
  FROM [MIKRAN].[dbo].[pd_Uzytkownik] 
  WHERE uz_Status > 0
  ORDER BY [uz_Nazwisko], [uz_Imie]";

                            Console.WriteLine("[SubiektSettings] ========================================");
                            Console.WriteLine("[SubiektSettings] Pobieranie listy użytkowników z MSSQL...");
                            Console.WriteLine("[SubiektSettings] Zapytanie SQL:");
                            Console.WriteLine($"[SubiektSettings] {sqlQuery}");
                            Console.WriteLine("[SubiektSettings] ========================================");

                            using (var command = new SqlCommand(sqlQuery, connection))
                            {
                                using (var reader = command.ExecuteReader())
                                {
                                    int rowCount = 0;
                                    while (reader.Read())
                                    {
                                        rowCount++;
                                        
                                        // Bezpieczne odczytywanie wartości z obsługą różnych typów
                                        int uzId = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
                                        string uzNazwisko = reader.IsDBNull(1) ? "" : reader.GetValue(1)?.ToString() ?? "";
                                        string uzImie = reader.IsDBNull(2) ? "" : reader.GetValue(2)?.ToString() ?? "";
                                        
                                        // Status może być int lub boolean - obsłuż oba przypadki
                                        object statusValue = reader.GetValue(3);
                                        string uzStatus = "";
                                        if (!reader.IsDBNull(3))
                                        {
                                            if (statusValue is bool boolStatus)
                                            {
                                                uzStatus = boolStatus ? "1" : "0";
                                            }
                                            else
                                            {
                                                uzStatus = Convert.ToString(statusValue) ?? "";
                                            }
                                        }

                                        // Format: "Nazwisko Imię"
                                        string displayName = $"{uzNazwisko} {uzImie}".Trim();
                                        
                                        // Dla logowania używamy również "Nazwisko Imię"
                                        string userName = displayName;

                                        var userItem = new UserItem
                                        {
                                            Id = uzId,
                                            UserName = userName,
                                            DisplayName = displayName
                                        };
                                        users.Add(userItem);

                                        Console.WriteLine($"[SubiektSettings] Użytkownik {rowCount}:");
                                        Console.WriteLine($"[SubiektSettings]   ID: {uzId}");
                                        Console.WriteLine($"[SubiektSettings]   Nazwisko: {uzNazwisko}");
                                        Console.WriteLine($"[SubiektSettings]   Imię: {uzImie}");
                                        Console.WriteLine($"[SubiektSettings]   Status: {uzStatus}");
                                        Console.WriteLine($"[SubiektSettings]   Wyświetlana nazwa: {displayName}");
                                        Console.WriteLine("[SubiektSettings] ---");
                                    }

                                    Console.WriteLine($"[SubiektSettings] ========================================");
                                    Console.WriteLine($"[SubiektSettings] Znaleziono {rowCount} użytkowników");
                                    Console.WriteLine("[SubiektSettings] ========================================");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SubiektSettings] Błąd podczas pobierania użytkowników: {ex.Message}");
                        throw;
                    }
                    
                    return users;
                });

                // Zaktualizuj ComboBox na wątku UI
                Dispatcher.Invoke(() =>
                {
                    _users.Clear();
                    
                    // Użytkownicy są już posortowani w zapytaniu SQL
                    foreach (var user in usersList)
                    {
                        _users.Add(user);
                    }
                    
                    // Jeśli istnieje zapisany użytkownik, ustaw go jako wybrany
                    string savedUser = _currentConfig.User ?? "";
                    if (!string.IsNullOrEmpty(savedUser))
                    {
                        UserComboBox.Text = savedUser;
                    }
                });

                MessageBox.Show(
                    $"Lista użytkowników została pobrana ({usersList.Count} użytkowników).\n\nWybierz użytkownika z listy powyżej.",
                    "Sukces",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (SqlException sqlEx)
            {
                string errorMessage = "Nie udało się pobrać listy użytkowników.\n\n";
                errorMessage += $"Błąd: {sqlEx.Message}";
                
                Console.WriteLine($"[SubiektSettings] Błąd SQL: {sqlEx.Message}");

                MessageBox.Show(errorMessage, "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Nie udało się pobrać listy użytkowników.\n\nBłąd: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                
                Console.WriteLine($"[SubiektSettings] Błąd: {ex.Message}");
            }
            finally
            {
                // Przywróć przycisk i kursor
                LoadUsersButton.IsEnabled = true;
                LoadUsersButton.Content = "📋 Pobierz listę użytkowników";
                Mouse.OverrideCursor = null;
            }
        }
    }
}

