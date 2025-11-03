using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Text.Json;
using System.IO;
using Gryzak.Views;
using Microsoft.Data.SqlClient;
using static Gryzak.Services.Logger;

namespace Gryzak.Services
{
    public class SubiektService
    {
        // Stała dla typu dokumentu ZK (zamówienie od klienta)
        private const int gtaSubiektDokumentZK = unchecked((int)0xFFFFFFF8); // -8
        
        // Cache dla obiektu GT i Subiekta - aby nie uruchamiać od nowa za każdym razem
        private static dynamic? _cachedGt = null;
        private static dynamic? _cachedSubiekt = null;
        
        // Cache dla mapy kodów VIES - aby nie wczytywać od nowa za każdym razem
        private static Dictionary<string, string>? _viesMapCache = null;

        private readonly ConfigService _configService;

        public SubiektService()
        {
            _configService = new ConfigService();
        }
        
        private class CountryInfo
        {
            public string name_pl { get; set; } = "";
            public string name_en { get; set; } = "";
            public string code { get; set; } = "";
            public string? code_vies { get; set; }
        }
        
        /// <summary>
        /// Wczytuje mapę kodów VIES z pliku kraje_iso2.json
        /// </summary>
        private static Dictionary<string, string> LoadViesMap()
        {
            if (_viesMapCache != null)
            {
                return _viesMapCache;
            }
            
            var viesMap = new Dictionary<string, string>();
            
            try
            {
                var jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "kraje_iso2.json");
                if (File.Exists(jsonPath))
                {
                    var jsonContent = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                    var countries = JsonSerializer.Deserialize<List<CountryInfo>>(jsonContent);
                    
                    if (countries != null)
                    {
                        foreach (var country in countries)
                        {
                            if (!string.IsNullOrWhiteSpace(country.code) && !string.IsNullOrWhiteSpace(country.code_vies))
                            {
                                viesMap[country.code.ToUpperInvariant()] = country.code_vies;
                            }
                        }
                        Info($"Wczytano {viesMap.Count} kodów VIES z pliku kraje_iso2.json", "SubiektService");
                    }
                }
            }
            catch (Exception ex)
            {
                Error($"Błąd podczas wczytywania mapy VIES: {ex.Message}", "SubiektService");
            }
            
            _viesMapCache = viesMap;
            return viesMap;
        }

        public bool CzyInstancjaAktywna()
        {
            return _cachedSubiekt != null && _cachedGt != null;
        }
        
        /// <summary>
        /// Wyszukuje kontrahenta przez zapytanie SQL do bazy MSSQL
        /// </summary>
        private int? WyszukajKontrahentaPrzezSQL(string? nip, string? email = null, string? customerName = null, string? phone = null, string? company = null, string? address = null, string? address1 = null, string? address2 = null, string? postcode = null, string? city = null, string? country = null, string? isoCode2 = null)
        {
            return WyszukajKontrahentaPrzezSQLInternal(nip, email, customerName, phone, company, address, address1, address2, postcode, city, country, isoCode2, false);
        }
        
        /// <summary>
        /// Wewnętrzna metoda do wyszukiwania kontrahenta przez SQL z obsługą rekurencji
        /// </summary>
        private int? WyszukajKontrahentaPrzezSQLInternal(string? nip, string? email = null, string? customerName = null, string? phone = null, string? company = null, string? address = null, string? address1 = null, string? address2 = null, string? postcode = null, string? city = null, string? country = null, string? isoCode2 = null, bool isRecursive = false)
        {
            int? selectedId = null;
            
            try
            {
                // Wczytaj konfigurację serwera MSSQL
                var subiektConfig = _configService.LoadSubiektConfig();
                string serverAddress = subiektConfig.ServerAddress ?? "";
                string username = subiektConfig.ServerUsername ?? "";
                string password = subiektConfig.ServerPassword ?? "";
                
                if (string.IsNullOrWhiteSpace(serverAddress))
                {
                    Debug("Brak adresu serwera MSSQL w konfiguracji - pomijam wyszukiwanie przez SQL.", "SubiektService");
                    return null;
                }
                
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
                
                // Wykonaj zapytanie SQL
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    
                    // Jeśli nie ma emaila ani nazwy klienta, nie wykonuj zapytania
                    if (string.IsNullOrWhiteSpace(email) && string.IsNullOrWhiteSpace(customerName))
                    {
                        Debug("Brak adresu email i nazwy klienta - pomijam wyszukiwanie przez SQL.", "SubiektService");
                        return null;
                    }
                    
                    // Buduj zapytanie SQL dynamicznie w zależności od dostępnych danych
                    string sqlQuery = @"
SELECT TOP(20)
    A.adr_TypAdresu,
    A.adr_Adres,
    A.adr_Nazwa,
    A.adr_NazwaPelna,
    A.adr_NIP,
    A.adr_Ulica,
    A.adr_Miejscowosc,
    A.adr_Kod,
    K.kh_Id,
    K.kh_Symbol,
    K.kh_EMail
FROM
    dbo.adr__Ewid AS A
INNER JOIN
    dbo.kh__Kontrahent AS K ON A.adr_IdObiektu = K.kh_Id
WHERE A.adr_TypAdresu = 1
AND (
";
                    
                    // Lista warunków do dodania
                    var conditions = new List<string>();
                    
                    // Warunek dla pełnej nazwy (imię i nazwisko)
                    if (!string.IsNullOrWhiteSpace(customerName))
                    {
                        // Podziel nazwę na słowa (imię i nazwisko)
                        var nameParts = customerName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        if (nameParts.Length > 0)
                        {
                            // Dodaj warunek dla każdego słowa z nazwy
                            var nameConditions = new List<string>();
                            foreach (var part in nameParts)
                            {
                                nameConditions.Add($"A.adr_NazwaPelna LIKE @NamePart{nameConditions.Count}");
                            }
                            conditions.Add($"({string.Join(" AND ", nameConditions)})");
                        }
                    }
                    
                    // Warunek dla adresu email
                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        conditions.Add("K.kh_EMail = @Email");
                    }
                    
                    sqlQuery += string.Join("\n        OR \n        ", conditions);
                    sqlQuery += "\n    )\n";
                    
                    Debug($"Wykonuję zapytanie SQL do wyszukania kontrahenta (email: {email}, customerName: {customerName}):", "SubiektService");
                    
                    using (var command = new SqlCommand(sqlQuery, connection))
                    {
                        // Dodaj parametr email do zapytania SQL (zabezpieczenie przed SQL injection)
                        if (!string.IsNullOrWhiteSpace(email))
                        {
                            command.Parameters.AddWithValue("@Email", email);
                        }
                        
                        // Dodaj parametry dla każdego słowa z nazwy (imię i nazwisko)
                        if (!string.IsNullOrWhiteSpace(customerName))
                        {
                            var nameParts = customerName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            for (int i = 0; i < nameParts.Length; i++)
                            {
                                command.Parameters.AddWithValue($"@NamePart{i}", $"%{nameParts[i]}%");
                            }
                        }
                        
                        // Loguj finalne zapytanie z parametrami
                        string loggedQuery = sqlQuery;
                        if (!string.IsNullOrWhiteSpace(email))
                        {
                            loggedQuery = loggedQuery.Replace("@Email", $"'{email}'");
                        }
                        if (!string.IsNullOrWhiteSpace(customerName))
                        {
                            var nameParts = customerName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            for (int i = 0; i < nameParts.Length; i++)
                            {
                                loggedQuery = loggedQuery.Replace($"@NamePart{i}", $"'%{nameParts[i]}%'");
                            }
                        }
                        Info("{loggedQuery}", "SubiektService");
                        
                        
                        using (var reader = command.ExecuteReader())
                        {
                            var kontrahenci = new ObservableCollection<KontrahentItem>();
                            
                            while (reader.Read())
                            {
                                var kontrahent = new KontrahentItem
                                {
                                    Id = reader.IsDBNull(8) ? 0 : Convert.ToInt32(reader.GetValue(8)), // kh_Id
                                    Symbol = reader.IsDBNull(9) ? "" : reader.GetValue(9)?.ToString() ?? "",
                                    NazwaPelna = reader.IsDBNull(3) ? "" : reader.GetValue(3)?.ToString() ?? "",
                                    Email = reader.IsDBNull(10) ? "" : reader.GetValue(10)?.ToString() ?? "",
                                    NIP = reader.IsDBNull(4) ? "" : reader.GetValue(4)?.ToString() ?? "",
                                    Adres = reader.IsDBNull(1) ? "" : reader.GetValue(1)?.ToString() ?? "",
                                    Miejscowosc = reader.IsDBNull(6) ? "" : reader.GetValue(6)?.ToString() ?? "",
                                    Kod = reader.IsDBNull(7) ? "" : reader.GetValue(7)?.ToString() ?? ""
                                };
                                
                                kontrahenci.Add(kontrahent);
                            }
                            
                            // Zawsze pokaż dialog wyboru, nawet jeśli nie znaleziono żadnych kontrahentów
                            // Użytkownik może zdecydować czy wybrać kontrahenta, dodać nowego, czy kontynuować bez kontrahenta
                            if (kontrahenci.Count > 0)
                            {
                                Info("Znaleziono {kontrahenci.Count} kontrahentów przez SQL.", "SubiektService");
                            }
                            else
                            {
                                Info("Nie znaleziono kontrahentów przez SQL - wyświetlam dialog z pustą listą.", "SubiektService");
                            }
                            
                            Info("Wyświetlam dialog wyboru kontrahenta ({kontrahenci.Count} wyników)...", "SubiektService");
                            
                            // Użyj synchronicznego Invoke, aby upewnić się że dialog jest wyświetlony
                            bool shouldAddNew = false;
                            bool shouldCancel = false;
                            Application.Current?.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    Info("Tworzę dialog SelectKontrahentDialog...", "SubiektService");
                                    var dialog = new SelectKontrahentDialog(kontrahenci, customerName, email, phone, company, nip, address);
                                    
                                    // Ustaw właściciela dialogu PO utworzeniu okna ale PRZED ShowDialog
                                    // W WPF można ustawić Owner tylko jeśli okno główne jest już wyświetlone
                                    bool hasOwner = false;
                                    try
                                    {
                                        if (Application.Current?.MainWindow != null && Application.Current.MainWindow.IsLoaded)
                                        {
                                            dialog.Owner = Application.Current.MainWindow;
                                            hasOwner = true;
                                        }
                                    }
                                    catch (Exception ownerEx)
                                    {
                                        Warning($"Nie można ustawić Owner dla dialogu: {ownerEx.Message}", "SubiektService");
                                        // Kontynuuj bez Owner - okno będzie wycentrowane na ekranie
                                        hasOwner = false;
                                    }
                                    
                                    // Ustaw lokalizację okna PRZED wyświetleniem (ale PO ustawieniu Owner)
                                    dialog.WindowStartupLocation = hasOwner ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
                                    
                                    // Ustaw dialog na wierzchu, aby był widoczny
                                    dialog.Topmost = true;
                                    
                                    Info("Wyświetlam dialog ShowDialog()...", "SubiektService");
                                    bool? result = dialog.ShowDialog();
                                    
                                    // Wyłącz Topmost po zamknięciu dialogu
                                    dialog.Topmost = false;
                                    Info("Dialog ShowDialog() zakończył się z wynikiem: {result}", "SubiektService");
                                    
                                    // Sprawdź najpierw specjalne akcje (kolejność ma znaczenie!)
                                    if (dialog.ShouldOpenEmpty)
                                    {
                                        Info("Użytkownik wybrał 'Pusty' - ZK zostanie otwarte bez kontrahenta.", "SubiektService");
                                        selectedId = null; // Explicitnie ustaw na null aby otworzyć ZK bez kontrahenta
                                    }
                                    else if (dialog.ShouldAddNew)
                                    {
                                        Info("Użytkownik chce dodać nowego kontrahenta - otwieram okno Subiekta GT", "SubiektService");
                                        shouldAddNew = true;
                                    }
                                    else if (result == true && dialog.SelectedKontrahent != null)
                                    {
                                        selectedId = dialog.SelectedKontrahent.Id;
                                        Info("Użytkownik wybrał kontrahenta: ID={dialog.SelectedKontrahent.Id}, Symbol={dialog.SelectedKontrahent.Symbol}", "SubiektService");
                                    }
                                    else
                                    {
                                        // Użytkownik kliknął Anuluj lub zamknął okno - nie otwieraj ZK
                                        Info("Użytkownik anulował wybór kontrahenta - dialog zostanie zamknięty bez otwierania ZK.", "SubiektService");
                                        shouldCancel = true; // Anuluj całkowicie - nie otwieraj ZK
                                    }
                                }
                                catch (Exception dialogEx)
                                {
                                    Error(dialogEx, "SubiektService", "Błąd podczas wyświetlania dialogu wyboru kontrahenta");
                                }
                            }, System.Windows.Threading.DispatcherPriority.Normal);
                            
                            // Jeśli użytkownik anulował, zakończ bez otwierania ZK
                            // Używamy -1 jako specjalnej wartości oznaczającej anulowanie
                            if (shouldCancel)
                            {
                                Info("Anulowano otwieranie ZK - zwracam -1 (specjalna wartość dla anulowania)", "SubiektService");
                                return -1;
                            }
                            
                            // Jeśli użytkownik chce dodać nowego kontrahenta, otwórz okno Subiekta GT
                            if (shouldAddNew)
                            {
                                Info("Otwieram okno dodawania kontrahenta...", "SubiektService");
                                try
                                {
                                    // Przekaż dane z API do metody DodajKontrahenta
                                    DodajKontrahenta(customerName, nip, company, email, phone, address1, address2, postcode, city, country, isoCode2);
                                }
                                catch (Exception addEx)
                                {
                                    Error(addEx, "SubiektService", "Błąd podczas otwierania okna dodawania kontrahenta");
                                }
                                
                                // Po zamknięciu okna Subiekta GT, wywołaj rekurencyjnie wyszukiwanie, aby ponownie pokazać dialog
                                Info("Okno Subiekta GT zostało zamknięte - ponownie wyświetlam dialog wyboru kontrahenta...", "SubiektService");
                                return WyszukajKontrahentaPrzezSQLInternal(nip, email, customerName, phone, company, address, address1, address2, postcode, city, country, isoCode2, true);
                            }
                            
                            Debug($"WyszukajKontrahentaPrzezSQL zwraca: {selectedId?.ToString() ?? "null"}", "SubiektService");
                            return selectedId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas wyszukiwania kontrahenta przez SQL");
                return null;
            }
        }
        
        /// <summary>
        /// Otwiera okno Subiekta GT do dodawania nowego kontrahenta
        /// </summary>
        public void DodajKontrahenta()
        {
            DodajKontrahenta(null, null, null, null, null, null, null, null, null, null, null);
        }
        
        /// <summary>
        /// Otwiera okno Subiekta GT do dodawania nowego kontrahenta z wypełnionymi danymi z API
        /// </summary>
        public void DodajKontrahenta(string? customerName = null, string? nip = null, string? company = null, string? email = null, string? phone = null, string? address1 = null, string? address2 = null, string? postcode = null, string? city = null, string? country = null, string? isoCode2 = null)
        {
            try
            {
                Info("Próba otwarcia okna dodawania kontrahenta...", "SubiektService");
                
                dynamic? gt = null;
                dynamic? subiekt = null;
                
                // Sprawdź czy mamy już uruchomioną instancję w cache
                if (_cachedSubiekt != null && _cachedGt != null)
                {
                    subiekt = _cachedSubiekt;
                    gt = _cachedGt;
                    Info("Używam istniejącej instancji Subiekta GT z cache.", "SubiektService");
                    PowiadomOZmianieInstancji(true);
                }
                
                // Jeśli nie ma działającej instancji, utwórz nową
                if (subiekt == null)
                {
                    Info("Uruchamiam nową instancję Subiekta GT...", "SubiektService");
                    
                    // Utworz obiekt GT (COM)
                    Type? gtType = Type.GetTypeFromProgID("InsERT.gt");
                    if (gtType == null)
                    {
                        Info("BŁĄD: Nie można załadować typu COM 'InsERT.gt'.", "SubiektService");
                        MessageBox.Show(
                            "Nie można połączyć się z Subiektem GT.\n\nUpewnij się, że:\n- Sfera dla Subiekta GT jest zainstalowana\n- Subiekt GT jest zainstalowany",
                            "Błąd połączenia",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                    
                    gt = Activator.CreateInstance(gtType);
                    if (gt == null)
                    {
                        Info("BŁĄD: Nie można utworzyć instancji obiektu GT.", "SubiektService");
                        return;
                    }
                    
                    // Ustaw parametry połączenia i logowania
                    try
                    {
                        // Wczytaj konfigurację przed ustawieniem parametrów
                        var subiektConfig = _configService.LoadSubiektConfig();
                        
                        // Sprawdź czy wszystkie wymagane parametry są ustawione
                        if (string.IsNullOrWhiteSpace(subiektConfig.ServerAddress))
                        {
                            Error("Brak adresu serwera w konfiguracji Subiekt GT.", "SubiektService");
                            MessageBox.Show(
                                "Brak adresu serwera w konfiguracji Subiekt GT.\n\nProszę skonfigurować połączenie w ustawieniach.",
                                "Brak konfiguracji",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Warning);
                            return;
                        }
                        
                        if (string.IsNullOrWhiteSpace(subiektConfig.DatabaseName))
                        {
                            Error("Brak nazwy bazy danych w konfiguracji Subiekt GT.", "SubiektService");
                            MessageBox.Show(
                                "Brak nazwy bazy danych w konfiguracji Subiekt GT.\n\nProszę skonfigurować połączenie w ustawieniach.",
                                "Brak konfiguracji",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Warning);
                            return;
                        }
                        
                        // Ustaw wszystkie wymagane parametry GT
                        gt.Produkt = 1; // gtaProduktSubiekt = 1
                        gt.Serwer = subiektConfig.ServerAddress;
                        gt.Baza = subiektConfig.DatabaseName;
                        gt.Autentykacja = 2; // gtaAutentykacjaMieszana = 2
                        
                        // Ustaw dane użytkownika bazy danych (jeśli podano)
                        if (!string.IsNullOrWhiteSpace(subiektConfig.ServerUsername))
                        {
                            gt.Uzytkownik = subiektConfig.ServerUsername;
                            gt.UzytkownikHaslo = subiektConfig.ServerPassword ?? "";
                        }
                        
                        // Ustaw operatora Subiekta
                        gt.Operator = subiektConfig.User ?? "Szef";
                        gt.OperatorHaslo = subiektConfig.Password ?? "";
                        
                        Info($"Ustawiono parametry GT - Serwer: {subiektConfig.ServerAddress}, Baza: {subiektConfig.DatabaseName}, Operator: {subiektConfig.User}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "SubiektService", "Błąd podczas konfiguracji GT");
                        MessageBox.Show(
                            $"Błąd podczas konfiguracji parametrów GT:\n\n{ex.Message}\n\nProszę sprawdzić konfigurację.",
                            "Błąd konfiguracji",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                    
                    // Uruchom Subiekta GT
                    try
                    {
                        subiekt = gt.Uruchom(2, 4); // Dopasuj operatora + uruchom w tle
                        if (subiekt == null)
                        {
                            Info("BŁĄD: Uruchomienie Subiekta GT zwróciło null.", "SubiektService");
                            return;
                        }
                        Info("Subiekt GT uruchomiony.", "SubiektService");
                        
                        // Zapisz w cache dla następnych użyć
                        _cachedGt = gt;
                        _cachedSubiekt = subiekt;
                        
                        // Powiadom o aktywności instancji
                        PowiadomOZmianieInstancji(true);
                        
                        // Ustaw główne okno jako niewidoczne
                        try
                        {
                            subiekt.Okno.Widoczne = false;
                            Info("Główne okno Subiekta GT ustawione jako niewidoczne.", "SubiektService");
                        }
                        catch (Exception oknoEx)
                        {
                            Warning($"Nie można ustawić głównego okna jako niewidoczne: {oknoEx.Message}", "SubiektService");
                        }
                    }
                    catch (COMException comEx)
                    {
                        Error($"Błąd COM: {comEx.Message}", "SubiektService");
                        _cachedSubiekt = null;
                        _cachedGt = null;
                        MessageBox.Show(
                            $"Błąd podczas uruchamiania Subiekta GT:\n\n{comEx.Message}\n\nUpewnij się, że Subiekt GT jest zainstalowany i Sfera jest aktywowana.",
                            "Błąd uruchamiania",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "SubiektService");
                        _cachedSubiekt = null;
                        _cachedGt = null;
                        MessageBox.Show(
                            $"Błąd podczas uruchamiania Subiekta GT:\n\n{ex.Message}",
                            "Błąd",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                }
                
                // Dodaj nowego kontrahenta
                try
                {
                    Info("Próba dodania nowego kontrahenta...", "SubiektService");
                    dynamic kontrahenci = subiekt.Kontrahenci;
                    if (kontrahenci != null)
                    {
                        dynamic nowyKontrahent = kontrahenci.Dodaj();
                        if (nowyKontrahent != null)
                        {
                            // Jeśli dane z API są dostępne, wypełnij pola kontrahenta
                            if (customerName != null || nip != null || company != null || email != null || phone != null || address1 != null || postcode != null || city != null)
                            {
                                WypelnijDaneKontrahenta(nowyKontrahent, customerName, nip, company, email, phone, address1, address2, postcode, city, country, isoCode2);
                            }
                            
                            // Wyświetl okno kartotekowe kontrahenta
                            nowyKontrahent.Wyswietl();
                            Info("Okno kartotekowe kontrahenta zostało otwarte.", "SubiektService");
                        }
                        else
                        {
                            Info("BŁĄD: Nie udało się utworzyć nowego kontrahenta.", "SubiektService");
                            MessageBox.Show(
                                "Nie udało się utworzyć nowego kontrahenta.",
                                "Błąd",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        Info("BŁĄD: Nie udało się uzyskać dostępu do kolekcji kontrahentów.", "SubiektService");
                        MessageBox.Show(
                            "Nie udało się uzyskać dostępu do kolekcji kontrahentów.",
                            "Błąd",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                    }
                }
                catch (Exception kontrEx)
                {
                    Error(kontrEx, "SubiektService", "Błąd podczas dodawania kontrahenta");
                    MessageBox.Show(
                        $"Błąd podczas dodawania kontrahenta:\n\n{kontrEx.Message}",
                        "Błąd",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas otwierania okna dodawania kontrahenta");
                MessageBox.Show(
                    $"Błąd podczas otwierania okna dodawania kontrahenta:\n\n{ex.Message}",
                    "Błąd",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Wyszukuje ID kraju w słowniku państw Subiekta GT po nazwie
        /// </summary>
        private int? WyszukajKrajWgNazwy(string countryName)
        {
            if (string.IsNullOrWhiteSpace(countryName))
                return null;
                
            try
            {
                var subiektConfig = _configService.LoadSubiektConfig();
                string serverAddress = subiektConfig.ServerAddress ?? "";
                string username = subiektConfig.ServerUsername ?? "";
                string password = subiektConfig.ServerPassword ?? "";
                
                if (string.IsNullOrWhiteSpace(serverAddress))
                {
                    Info("Brak adresu serwera MSSQL - pomijam wyszukiwanie kraju.", "SubiektService");
                    return null;
                }
                
                var builder = new SqlConnectionStringBuilder
                {
                    DataSource = serverAddress,
                    UserID = username,
                    Password = password,
                    ConnectTimeout = 10,
                    Encrypt = false
                };
                
                if (string.IsNullOrWhiteSpace(username))
                {
                    builder.IntegratedSecurity = true;
                }
                
                string connectionString = builder.ConnectionString;
                
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    
                    // Szukaj kraju w słowniku państw
                    string sqlQuery = @"
SELECT TOP(1) pa_Id
FROM dbo.sl_Panstwo
WHERE pa_Nazwa = @CountryName";
                    
                    using (var command = new SqlCommand(sqlQuery, connection))
                    {
                        command.Parameters.AddWithValue("@CountryName", countryName);
                        
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var countryId = reader.IsDBNull(0) ? null : (int?)Convert.ToInt32(reader.GetValue(0));
                                if (countryId.HasValue)
                                {
                                    Info("Znaleziono kraj '{countryName}' w słowniku: ID={countryId.Value}", "SubiektService");
                                    return countryId.Value;
                                }
                            }
                        }
                    }
                }
                
                Info("Nie znaleziono kraju '{countryName}' w słowniku państw.", "SubiektService");
                return null;
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas wyszukiwania kraju");
                return null;
            }
        }
        
        /// <summary>
        /// Sprawdza czy dokument z podanym numerem oryginalnym już istnieje w Subiekcie
        /// Zwraca dok_Id dokumentu jeśli istnieje, null w przeciwnym razie
        /// </summary>
        private int? PobierzIdIstniejacegoDokumentu(string numerOryginalny)
        {
            if (string.IsNullOrWhiteSpace(numerOryginalny))
            {
                return null;
            }
            
            try
            {
                var subiektConfig = _configService.LoadSubiektConfig();
                string serverAddress = subiektConfig.ServerAddress ?? "";
                string username = subiektConfig.ServerUsername ?? "";
                string password = subiektConfig.ServerPassword ?? "";
                
                if (string.IsNullOrWhiteSpace(serverAddress))
                {
                    Info("Brak adresu serwera MSSQL - pomijam sprawdzanie istnienia dokumentu.", "SubiektService");
                    return null;
                }
                
                var builder = new SqlConnectionStringBuilder
                {
                    DataSource = serverAddress,
                    UserID = username,
                    Password = password,
                    ConnectTimeout = 10,
                    Encrypt = false
                };
                
                if (string.IsNullOrWhiteSpace(username))
                {
                    builder.IntegratedSecurity = true;
                }
                
                string connectionString = builder.ConnectionString;
                
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    
                    string sqlQuery = @"
SELECT 
    [dok_Id],
    [dok_NrPelnyOryg]
FROM [MIKRAN].[dbo].[dok__Dokument]
WHERE dok_NrPelnyOryg = @NumerOryginalny";
                    
                    using (var command = new SqlCommand(sqlQuery, connection))
                    {
                        command.Parameters.AddWithValue("@NumerOryginalny", numerOryginalny);
                        
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var dokId = reader.IsDBNull(0) ? null : (int?)Convert.ToInt32(reader.GetValue(0));
                                var nrPelnyOryg = reader.IsDBNull(1) ? null : reader.GetString(1);
                                Info("Znaleziono istniejący dokument z numerem oryginalnym '{numerOryginalny}' (dok_Id={dokId})", "SubiektService");
                                return dokId;
                            }
                        }
                    }
                }
                
                Info("Nie znaleziono dokumentu z numerem oryginalnym '{numerOryginalny}'", "SubiektService");
                return null;
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas sprawdzania istnienia dokumentu");
                // W przypadku błędu zwracamy null, aby nie blokować otwierania ZK
                return null;
            }
        }
        
        /// <summary>
        /// Wypełnia dane kontrahenta w obiekcie Subiekta GT
        /// </summary>
        private void WypelnijDaneKontrahenta(dynamic kontrahent, string? customerName, string? nip, string? company, string? email, string? phone, string? address1, string? address2, string? postcode, string? city, string? country, string? isoCode2 = null)
        {
            try
            {
                Info("Wypełnianie danych kontrahenta...", "SubiektService");
                
                // Ustaw typ kontrahenta - 2 = gtaKontrahentTypOdbiorca (Odbiorca)
                try
                {
                    kontrahent.Typ = 2;
                    Info("Ustawiono Typ=2 (Odbiorca)", "SubiektService");
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się ustawić Typ: {ex.Message}", "SubiektService");
                }
                
                // Ustaw województwo na (brak) - gtaBrak = -2147483648
                try
                {
                    kontrahent.Wojewodztwo = unchecked((int)0x80000000);
                    Info("Ustawiono Wojewodztwo=(brak)", "SubiektService");
                }
                catch (Exception ex)
                {
                    Warning($"Nie udało się ustawić Wojewodztwo: {ex.Message}", "SubiektService");
                }
                
                // Zbuduj pełną nazwę: imię, nazwisko i firma (jeśli dostępne)
                string pelnaNazwa = "";
                if (!string.IsNullOrWhiteSpace(customerName))
                {
                    pelnaNazwa = customerName;
                }
                if (!string.IsNullOrWhiteSpace(company))
                {
                    if (!string.IsNullOrWhiteSpace(pelnaNazwa))
                    {
                        pelnaNazwa += ", " + company;
                    }
                    else
                    {
                        pelnaNazwa = company;
                    }
                }
                
                // Jeśli mamy jakąkolwiek nazwę, ustaw wszystkie pola
                if (!string.IsNullOrWhiteSpace(pelnaNazwa))
                {
                    // Symbol - Subiekt GT może generować to automatycznie lub wymaga unikalnej wartości
                    // Pomijamy ustawianie Symbol programowo, aby uniknąć błędów związanych z duplikatami
                    // try
                    // {
                    //     kontrahent.Symbol = pelnaNazwa;
                    //     Info("Ustawiono Symbol: {pelnaNazwa}", "SubiektService");
                    // }
                    // catch (Exception ex)
                    // {
                    //     Info("Nie udało się ustawić Symbol: {ex.Message}", "SubiektService");
                    // }
                    
                    // Nazwa
                    try
                    {
                        kontrahent.Nazwa = pelnaNazwa;
                        Info("Ustawiono Nazwa: {pelnaNazwa}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić Nazwa: {ex.Message}", "SubiektService");
                    }
                    
                    // Nazwa pełna
                    try
                    {
                        kontrahent.NazwaPelna = pelnaNazwa;
                        Info("Ustawiono NazwaPelna: {pelnaNazwa}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić NazwaPelna: {ex.Message}", "SubiektService");
                    }
                }
                
                // NIP - dodaj prefiks UE jeśli dostępny
                if (!string.IsNullOrWhiteSpace(nip))
                {
                    try
                    {
                        string nipZPrefiksem = nip;
                        
                        // Jeśli mamy kod ISO2, spróbuj dodać prefiks VIES
                        if (!string.IsNullOrWhiteSpace(isoCode2))
                        {
                            var viesMap = LoadViesMap();
                            if (viesMap.TryGetValue(isoCode2.ToUpperInvariant(), out var viesCode))
                            {
                                nipZPrefiksem = viesCode + nip;
                                Info("Dodaję prefiks VIES do NIP: {nipZPrefiksem}", "SubiektService");
                            }
                        }
                        
                        kontrahent.NIP = nipZPrefiksem;
                        Info("Ustawiono NIP: {nipZPrefiksem}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić NIP: {ex.Message}", "SubiektService");
                    }
                }
                
                // Ulica
                if (!string.IsNullOrWhiteSpace(address1))
                {
                    try
                    {
                        kontrahent.Ulica = address1;
                        Info("Ustawiono Ulica: {address1}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić Ulica: {ex.Message}", "SubiektService");
                    }
                }
                
                // Miejscowość
                if (!string.IsNullOrWhiteSpace(city))
                {
                    try
                    {
                        kontrahent.Miejscowosc = city;
                        Info("Ustawiono Miejscowosc: {city}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić Miejscowosc: {ex.Message}", "SubiektService");
                    }
                }
                
                // Kod pocztowy
                if (!string.IsNullOrWhiteSpace(postcode))
                {
                    try
                    {
                        kontrahent.KodPocztowy = postcode;
                        Info("Ustawiono KodPocztowy: {postcode}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić KodPocztowy: {ex.Message}", "SubiektService");
                    }
                }
                
                // Email
                if (!string.IsNullOrWhiteSpace(email) && email != "Brak email")
                {
                    try
                    {
                        dynamic emaile = kontrahent.Emaile;
                        if (emaile != null)
                        {
                            dynamic nowyEmail = emaile.Dodaj(email);
                            if (nowyEmail != null)
                            {
                                nowyEmail.Podstawowy = true;
                                Info("Ustawiono Email: {email}", "SubiektService");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić Email: {ex.Message}", "SubiektService");
                    }
                }
                
                // Telefon
                if (!string.IsNullOrWhiteSpace(phone) && phone != "Brak telefonu")
                {
                    try
                    {
                        dynamic telefony = kontrahent.Telefony;
                        if (telefony != null)
                        {
                            dynamic nowyTelefon = telefony.Dodaj(phone);
                            if (nowyTelefon != null)
                            {
                                nowyTelefon.Podstawowy = true;
                                Info("Ustawiono Telefon: {phone}", "SubiektService");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Warning($"Nie udało się ustawić Telefon: {ex.Message}", "SubiektService");
                    }
                }
                
                // Kraj - wyszukaj ID w słowniku państw i ustaw
                // UWAGA: Najpierw ustawiamy PodatnikVatUE, a potem Panstwo, aby event wywołał się poprawnie
                if (!string.IsNullOrWhiteSpace(country))
                {
                    try
                    {
                        // Najpierw ustaw PodatnikVatUE jeśli kraj jest w UE (ale nie dla Polski)
                        // Ustawiamy tylko jeśli NIP nie jest pusty i nie składa się tylko z prefiksu VIES
                        if (!string.IsNullOrWhiteSpace(nip) && !string.IsNullOrWhiteSpace(isoCode2))
                        {
                            var isoCodeUpper = isoCode2.ToUpperInvariant();
                            // Pomijamy Polskę - tylko kraje UE poza Polską
                            if (isoCodeUpper != "PL")
                            {
                                var viesMap = LoadViesMap();
                                if (viesMap.TryGetValue(isoCodeUpper, out var viesCode))
                                {
                                    // Sprawdź czy NIP nie składa się tylko z prefiksu VIES (nie ma dodatkowych cyfr)
                                    // Używamy raw NIP (bez prefiksu) aby sprawdzić czy jest rzeczywiście wypełniony
                                    bool nipJestPelny = nip.Length > viesCode.Length;
                                    
                                    if (nipJestPelny)
                                    {
                                        try
                                        {
                                            kontrahent.PodatnikVatUE = true;
                                            Info("Ustawiono PodatnikVatUE=True (kraj UE: {viesCode}, NIP: {nip})", "SubiektService");
                                        }
                                        catch (Exception ex)
                                        {
                                            Warning($"Nie udało się ustawić PodatnikVatUE: {ex.Message}", "SubiektService");
                                        }
                                    }
                                    else
                                    {
                                        Info("Pomijam ustawienie PodatnikVatUE - NIP jest pusty lub składa się tylko z prefiksu VIES ({viesCode})", "SubiektService");
                                    }
                                }
                            }
                            else
                            {
                                Info("Pomijam ustawienie PodatnikVatUE dla Polski (PL)", "SubiektService");
                            }
                        }
                        else if (string.IsNullOrWhiteSpace(nip))
                        {
                            Info("Pomijam ustawienie PodatnikVatUE - brak NIP", "SubiektService");
                        }
                        
                        // Teraz ustaw Panstwo (to powinno wywołać event w Subiekcie GT)
                        var countryId = WyszukajKrajWgNazwy(country);
                        if (countryId.HasValue)
                        {
                            kontrahent.Panstwo = countryId.Value;
                            Info("Ustawiono Panstwo: {country} (ID={countryId.Value})", "SubiektService");
                        }
                        else
                        {
                            Info("Nie znaleziono kraju '{country}' w słowniku - pomijam ustawienie Panstwo", "SubiektService");
                        }
                    }
                    catch (Exception ex)
                    {
                        Warning($"Błąd podczas ustawiania Panstwo: {ex.Message}", "SubiektService");
                    }
                }
                
                Info("Dane kontrahenta zostały wypełnione.", "SubiektService");
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas wypełniania danych kontrahenta");
            }
        }
        
        /// <summary>
        /// Zamyka aktywną instancję Subiekta GT i zwalnia licencję
        /// </summary>
        public void ZwolnijLicencje()
        {
            try
            {
                Info("Zamykanie instancji Subiekta GT i zwolnienie licencji...", "SubiektService");
                
                if (_cachedSubiekt != null)
                {
                    try
                    {
                        // Próbuj uzyskać dostęp do obiektu Aplikacja i zamknąć całą aplikację
                        try
                        {
                            dynamic aplikacja = _cachedSubiekt.Aplikacja;
                            if (aplikacja != null)
                            {
                                bool zakonczWynik = aplikacja.Zakoncz();
                                Info("Wywołano Zakoncz() na obiekcie Aplikacja. Wynik: {zakonczWynik}", "SubiektService");
                                
                                // Poczekaj chwilę na zamknięcie procesu
                                System.Threading.Thread.Sleep(500);
                            }
                        }
                        catch (Exception aplikacjaEx)
                        {
                            Warning($"Nie udało się zamknąć przez Aplikacja.Zakoncz(): {aplikacjaEx.Message}", "SubiektService");
                            // Próbuj alternatywną metodę - Zakoncz na obiekcie GT
                            try
                            {
                                if (_cachedGt != null)
                                {
                                    // Spróbuj wywołać Zakoncz na obiekcie GT (jeśli dostępne)
                                    try
                                    {
                                        dynamic aplikacjaGT = _cachedGt.Aplikacja;
                                        if (aplikacjaGT != null)
                                        {
                                            bool zakonczWynik = aplikacjaGT.Zakoncz();
                                            Info("Wywołano Zakoncz() na obiekcie GT.Aplikacja. Wynik: {zakonczWynik}", "SubiektService");
                                            System.Threading.Thread.Sleep(500);
                                        }
                                    }
                                    catch
                                    {
                                        Info("Nie można uzyskać dostępu do GT.Aplikacja", "SubiektService");
                                    }
                                }
                            }
                            catch
                            {
                                // Ignoruj błędy
                            }
                        }
                    }
                    catch (Exception zamknijEx)
                    {
                        Warning($"Błąd podczas próby zamknięcia aplikacji: {zamknijEx.Message}", "SubiektService");
                    }
                    
                    // Zwolnij referencje COM obiektów
                    try
                    {
                        if (_cachedSubiekt != null)
                        {
                            Marshal.ReleaseComObject(_cachedSubiekt);
                            Info("Zwolniono referencję COM dla obiektu Subiekt.", "SubiektService");
                        }
                    }
                    catch (Exception releaseEx)
                    {
                        Warning($"Błąd podczas zwalniania obiektu Subiekt: {releaseEx.Message}", "SubiektService");
                    }
                }
                
                if (_cachedGt != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(_cachedGt);
                        Info("Zwolniono referencję COM dla obiektu GT.", "SubiektService");
                    }
                    catch (Exception releaseEx)
                    {
                        Warning($"Błąd podczas zwalniania obiektu GT: {releaseEx.Message}", "SubiektService");
                    }
                }
                
                // Wyczyść cache niezależnie od tego czy zamknięcie się powiodło
                _cachedSubiekt = null;
                _cachedGt = null;
                
                // Wymuś garbage collection aby zwolnić wszystkie referencje COM
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Powiadom o zmianie statusu
                PowiadomOZmianieInstancji(false);
                
                Info("Instancja Subiekta GT została zamknięta, licencja zwolniona.", "SubiektService");
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektService", "Błąd podczas zamykania instancji");
                // Mimo błędu wyczyść cache
                _cachedSubiekt = null;
                _cachedGt = null;
                PowiadomOZmianieInstancji(false);
            }
        }
        
        // Event do powiadomienia o zmianie statusu instancji
        public static event EventHandler<bool>? InstancjaZmieniona;

        private static void PowiadomOZmianieInstancji(bool aktywna)
        {
            InstancjaZmieniona?.Invoke(null, aktywna);
        }

        public void OtworzOknoZK(string? nip = null, System.Collections.Generic.IEnumerable<Gryzak.Models.Product>? items = null, double? couponAmount = null, double? subTotal = null, string? couponTitle = null, string? orderId = null, double? handlingAmount = null, double? shippingAmount = null, string? currency = null, double? codFeeAmount = null, string? orderTotal = null, double? glsAmount = null, string? email = null, string? customerName = null, string? phone = null, string? company = null, string? address = null, string? address1 = null, string? address2 = null, string? postcode = null, string? city = null, string? country = null, string? isoCode2 = null)
        {
            try
            {
                Info($"Próba otwarcia okna ZK{(nip != null ? $" z kontrahentem o NIP: {nip}" : " bez kontrahenta")}{(email != null ? $" (email: {email})" : "")}...", "SubiektService");

                dynamic? gt = null;
                dynamic? subiekt = null;

                // Sprawdź czy mamy już uruchomioną instancję w cache
                // Używamy prostej weryfikacji - jeśli obiekt jest w cache, zakładamy że działa
                // Jeśli nie działa, catch podczas użycia złapie wyjątek i wtedy wyczyścimy cache
                if (_cachedSubiekt != null && _cachedGt != null)
                {
                    subiekt = _cachedSubiekt;
                    gt = _cachedGt;
                    Info("Używam istniejącej instancji Subiekta GT z cache.", "SubiektService");
                    // Upewnij się, że status jest aktywny (może nie zostać ustawiony przy starcie aplikacji)
                    PowiadomOZmianieInstancji(true);
                }

                // Jeśli nie ma działającej instancji, utwórz nową
                if (subiekt == null)
                {
                    Info("Uruchamiam nową instancję Subiekta GT...", "SubiektService");

                    // Utworz obiekt GT (COM)
                    Type? gtType = Type.GetTypeFromProgID("InsERT.gt");
                    if (gtType == null)
                    {
                        Info("BŁĄD: Nie można załadować typu COM 'InsERT.gt'. Upewnij się, że Sfera jest zainstalowana.", "SubiektService");
                        MessageBox.Show(
                            "Nie można połączyć się z Subiektem GT.\n\nUpewnij się, że:\n- Sfera dla Subiekta GT jest zainstalowana\n- Subiekt GT jest zainstalowany",
                            "Błąd połączenia",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }

                    gt = Activator.CreateInstance(gtType);
                    if (gt == null)
                    {
                        Info("BŁĄD: Nie można utworzyć instancji obiektu GT.", "SubiektService");
                        return;
                    }

                    // Ustaw parametry połączenia i logowania
                    try
                    {
                        // Wczytaj konfigurację przed ustawieniem parametrów
                        var subiektConfig = _configService.LoadSubiektConfig();
                        
                        // Sprawdź czy wszystkie wymagane parametry są ustawione
                        if (string.IsNullOrWhiteSpace(subiektConfig.ServerAddress))
                        {
                            Error("Brak adresu serwera w konfiguracji Subiekt GT.", "SubiektService");
                            MessageBox.Show(
                                "Brak adresu serwera w konfiguracji Subiekt GT.\n\nProszę skonfigurować połączenie w ustawieniach.",
                                "Brak konfiguracji",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Warning);
                            return;
                        }
                        
                        if (string.IsNullOrWhiteSpace(subiektConfig.DatabaseName))
                        {
                            Error("Brak nazwy bazy danych w konfiguracji Subiekt GT.", "SubiektService");
                            MessageBox.Show(
                                "Brak nazwy bazy danych w konfiguracji Subiekt GT.\n\nProszę skonfigurować połączenie w ustawieniach.",
                                "Brak konfiguracji",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Warning);
                            return;
                        }
                        
                        // Ustaw wszystkie wymagane parametry GT
                        gt.Produkt = 1; // gtaProduktSubiekt = 1
                        gt.Serwer = subiektConfig.ServerAddress;
                        gt.Baza = subiektConfig.DatabaseName;
                        gt.Autentykacja = 2; // gtaAutentykacjaMieszana = 2
                        
                        // Ustaw dane użytkownika bazy danych (jeśli podano)
                        if (!string.IsNullOrWhiteSpace(subiektConfig.ServerUsername))
                        {
                            gt.Uzytkownik = subiektConfig.ServerUsername;
                            gt.UzytkownikHaslo = subiektConfig.ServerPassword ?? "";
                        }
                        
                        // Ustaw operatora Subiekta
                        gt.Operator = subiektConfig.User ?? "Szef";
                        gt.OperatorHaslo = subiektConfig.Password ?? "";
                        
                        Info($"Ustawiono parametry GT - Serwer: {subiektConfig.ServerAddress}, Baza: {subiektConfig.DatabaseName}, Operator: {subiektConfig.User}", "SubiektService");
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "SubiektService", "Błąd podczas konfiguracji GT");
                        MessageBox.Show(
                            $"Błąd podczas konfiguracji parametrów GT:\n\n{ex.Message}\n\nProszę sprawdzić konfigurację.",
                            "Błąd konfiguracji",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }

                    // Uruchom Subiekta GT w tle (bez interfejsu użytkownika)
                    // gtaUruchomWTle (0x4) - aplikacja działa w tle, nie otwiera się jej okno
                    try
                    {
                        // gtaUruchomDopasujOperatora = 2, gtaUruchomWTle = 4
                        // To uruchomi Subiekta w tle bez pokazywania jakichkolwiek okien (włącznie z logowaniem)
                        subiekt = gt.Uruchom(2, 4); // Dopasuj operatora + uruchom w tle
                        if (subiekt == null)
                        {
                            Info("BŁĄD: Uruchomienie Subiekta GT zwróciło null.", "SubiektService");
                            return;
                        }
                        Info("Subiekt GT uruchomiony w tle (bez interfejsu użytkownika).", "SubiektService");
                        

                        
                        // Zapisz w cache dla następnych użyć
                        _cachedGt = gt;
                        _cachedSubiekt = subiekt;
                        
                        // Powiadom o aktywności instancji
                        PowiadomOZmianieInstancji(true);
                        
                        // Upewnij się, że główne okno jest ukryte (dodatkowa ochrona)
                        try
                        {
                            subiekt.Okno.Widoczne = false;
                            Info("Główne okno Subiekta GT ustawione jako niewidoczne.", "SubiektService");
                        }
                        catch (Exception oknoEx)
                        {
                            Warning($"Nie można ustawić głównego okna jako niewidoczne: {oknoEx.Message}", "SubiektService");
                        }
                    }
                        catch (COMException comEx)
                    {
                        Error($"Błąd COM: {comEx.Message} (HRESULT: 0x{comEx.ErrorCode:X8})", "SubiektService");
                        // Wyczyść cache w przypadku błędu COM
                        _cachedSubiekt = null;
                        _cachedGt = null;
                        PowiadomOZmianieInstancji(false);
                        MessageBox.Show(
                            $"Błąd podczas uruchamiania Subiekta GT:\n\n{comEx.Message}\n\nUpewnij się, że Subiekt GT jest zainstalowany i Sfera jest aktywowana.",
                            "Błąd uruchamiania",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "SubiektService");
                        // Wyczyść cache w przypadku błędu
                        _cachedSubiekt = null;
                        _cachedGt = null;
                        MessageBox.Show(
                            $"Błąd podczas uruchamiania Subiekta GT:\n\n{ex.Message}",
                            "Błąd",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                        return;
                    }
                }

                // Sprawdź czy dokument z numerem oryginalnym już istnieje w Subiekcie
                // Jeśli istnieje, wyświetl ostrzeżenie i zapytaj czy otworzyć istniejący dokument
                if (!string.IsNullOrWhiteSpace(orderId) && subiekt != null)
                {
                    int? dokId = PobierzIdIstniejacegoDokumentu(orderId!); // orderId nie może być null w tym miejscu
                    if (dokId.HasValue)
                    {
                        Info("Dokument z numerem oryginalnym '{orderId}' już istnieje (dok_Id={dokId.Value}).", "SubiektService");
                        
                        // Wyświetl ostrzeżenie i zapytaj użytkownika
                        var result = MessageBox.Show(
                            $"Dokument z numerem oryginalnym '{orderId}' już istnieje w Subiekcie.\n\n" +
                            $"Czy chcesz otworzyć istniejący dokument do edycji?",
                            "Ostrzeżenie - Dokument już istnieje",
                            System.Windows.MessageBoxButton.YesNo,
                            System.Windows.MessageBoxImage.Warning);
                        
                        if (result == System.Windows.MessageBoxResult.No)
                        {
                            Info("Użytkownik anulował otwieranie dokumentu ZK - dokument z numerem '{orderId}' już istnieje.", "SubiektService");
                            return;
                        }
                        
                        // Użytkownik wybrał "Tak" - otwórz istniejący dokument
                        Info("Użytkownik zdecydował otworzyć istniejący dokument (dok_Id={dokId.Value}).", "SubiektService");
                        
                        try
                        {
                            // Wczytaj istniejący dokument po ID
                            // subiekt nie może być null w tym miejscu (sprawdzony w warunku if)
                            dynamic dokumenty = subiekt!.Dokumenty;
                            dynamic istniejacyDokument = dokumenty.Wczytaj(dokId.Value);
                            
                            if (istniejacyDokument != null)
                            {
                                // Otwórz dokument do edycji (bez uzupełniania danych)
                                istniejacyDokument.Wyswietl(false);
                                Info("Otwarto istniejący dokument ZK (dok_Id={dokId.Value}) do edycji.", "SubiektService");
                                
                                // Aktywuj okno dokumentu na wierzch
                                try
                                {
                                    subiekt!.Okno.Aktywuj(); // subiekt nie może być null (sprawdzony w warunku if)
                                    Info("Aktywowano okno dokumentu na wierzch.", "SubiektService");
                                }
                                catch (Exception aktywujEx)
                                {
                                    Warning($"Aktywacja okna dokumentu nie zadziałała: {aktywujEx.Message}", "SubiektService");
                                }
                                
                                // Zakończ funkcję - nie uzupełniaj żadnych danych
                                return;
                            }
                            else
                            {
                                Info("Nie udało się wczytać dokumentu o ID={dokId.Value}. Tworzę nowy dokument.", "SubiektService");
                            }
                        }
                        catch (Exception wczytajEx)
                        {
                            Error(wczytajEx, "SubiektService", "Błąd podczas wczytywania istniejącego dokumentu. Tworzę nowy dokument.");
                            // Kontynuuj normalną ścieżką tworzenia nowego dokumentu
                        }
                    }
                }

                // Dodaj dokument ZK i otwórz jego okno
                // Używamy zmiennej subiekt, która może pochodzić z cache lub być świeżo utworzona
                if (subiekt == null)
                {
                    Info("BŁĄD: subiekt jest null - nie można utworzyć dokumentu ZK.", "SubiektService");
                    return;
                }
                
                try
                {
                    Info("Próba dodania dokumentu ZK...", "SubiektService");
                    dynamic? zkDokument = null;

                    // Spróbuj użyć SuDokumentyManager.DodajZK() - to jest zalecana metoda
                    try
                    {
                        dynamic dokumentyManager = subiekt.SuDokumentyManager;
                        if (dokumentyManager != null)
                        {
                            zkDokument = dokumentyManager.DodajZK();
                            Info("Użyto SuDokumentyManager.DodajZK()", "SubiektService");
                        }
                    }
                    catch (Exception ex1)
                    {
                        Warning($"SuDokumentyManager.DodajZK() nie zadziałał: {ex1.Message}", "SubiektService");
                    }

                    // Fallback: użyj Dokumenty.DodajZK()
                    if (zkDokument == null)
                    {
                        try
                        {
                            dynamic dokumenty = subiekt!.Dokumenty; // subiekt sprawdzony wyżej
                            zkDokument = dokumenty.DodajZK();
                            Info("Użyto Dokumenty.DodajZK()", "SubiektService");
                        }
                        catch (Exception ex2)
                        {
                            Warning($"Dokumenty.DodajZK() nie zadziałał: {ex2.Message}", "SubiektService");
                        }
                    }

                    // Ostatni fallback: użyj Dokumenty.Dodaj() z parametrem
                    if (zkDokument == null)
                    {
                        try
                        {
                            dynamic dokumenty = subiekt!.Dokumenty; // subiekt sprawdzony wyżej
                            zkDokument = dokumenty.Dodaj(gtaSubiektDokumentZK);
                            Info("Użyto Dokumenty.Dodaj(gtaSubiektDokumentZK)", "SubiektService");
                        }
                        catch (Exception ex3)
                        {
                            Warning($"Dokumenty.Dodaj() nie zadziałał: {ex3.Message}", "SubiektService");
                        }
                    }

                    if (zkDokument != null)
                    {
                        Info("Dokument ZK został utworzony.", "SubiektService");
                        
                        // Wyszukaj kontrahenta po NIP i ustaw w dokumencie (tylko jeśli NIP został podany)
                        if (!string.IsNullOrWhiteSpace(nip))
                        {
                            try
                            {
                                // Pobierz menedżer kontrahentów
                                dynamic kontrahenciManager = subiekt!.KontrahenciManager; // subiekt sprawdzony wyżej
                                
                                // Wyszukaj kontrahenta po NIP - gtaKontrahentWgNip = 2
                                dynamic kontrahent = kontrahenciManager.WczytajKontrahentaWg(nip, 2);
                                
                                if (kontrahent != null)
                                {
                                    // Ustaw ID kontrahenta w dokumencie
                                    zkDokument.KontrahentId = kontrahent.Identyfikator;
                                    Info("Ustawiono kontrahenta o ID={kontrahent.Identyfikator} (NIP: {nip}) w dokumencie ZK.", "SubiektService");
                                }
                                else
                                {
                                    Info("Nie znaleziono kontrahenta o NIP: {nip}", "SubiektService");
                                    
                                    // Spróbuj wyszukać kontrahenta przez SQL po emailu
                                    Info("Próba wyszukania kontrahenta przez SQL...", "SubiektService");
                                    int? kontrahentId = WyszukajKontrahentaPrzezSQL(nip, email, customerName, phone, company, address, address1, address2, postcode, city, country, isoCode2);
                                    
                                    // Sprawdź czy użytkownik anulował (wartość -1 oznacza anulowanie)
                                    if (kontrahentId.HasValue && kontrahentId.Value == -1)
                                    {
                                        Info("Użytkownik anulował wybór kontrahenta - przerywam otwieranie ZK.", "SubiektService");
                                        return; // Przerwij otwieranie ZK
                                    }
                                    
                                    if (kontrahentId.HasValue && kontrahentId.Value > 0)
                                    {
                                        // Ustaw ID kontrahenta w dokumencie
                                        zkDokument.KontrahentId = kontrahentId.Value;
                                        Info("Ustawiono kontrahenta o ID={kontrahentId.Value} (znaleziony przez SQL) w dokumencie ZK.", "SubiektService");
                                    }
                                    else
                                    {
                                        Info("Nie znaleziono kontrahenta przez SQL - ZK zostanie otwarte bez kontrahenta.", "SubiektService");
                                    }
                                }
                            }
                            catch (Exception kontrEx)
                            {
                                Error(kontrEx, "SubiektService", "Błąd: Nie udało się wyszukać kontrahenta");
                                
                                // Próbuj wyszukać przez SQL nawet jeśli wystąpił błąd
                                try
                                {
                                    Info("Próba wyszukania kontrahenta przez SQL (po błędzie NIP)...", "SubiektService");
                                    int? kontrahentId = WyszukajKontrahentaPrzezSQL(nip, email, customerName, phone, company, address, address1, address2, postcode, city, country, isoCode2);
                                    
                                    // Sprawdź czy użytkownik anulował (wartość -1 oznacza anulowanie)
                                    if (kontrahentId.HasValue && kontrahentId.Value == -1)
                                    {
                                        Info("Użytkownik anulował wybór kontrahenta - przerywam otwieranie ZK.", "SubiektService");
                                        return; // Przerwij otwieranie ZK
                                    }
                                    
                                    if (kontrahentId.HasValue && kontrahentId.Value > 0 && zkDokument != null)
                                    {
                                        // Ustaw ID kontrahenta w dokumencie
                                        int id = kontrahentId!.Value; // HasValue jest true, więc Value nie jest null
                                        zkDokument!.KontrahentId = id; // Sprawdziliśmy != null wyżej
                                        Info("Ustawiono kontrahenta o ID={id} (znaleziony przez SQL po błędzie) w dokumencie ZK.", "SubiektService");
                                    }
                                    else
                                    {
                                        Info("Nie znaleziono kontrahenta przez SQL po błędzie - ZK zostanie otwarte bez kontrahenta.", "SubiektService");
                                    }
                                }
                                catch (Exception sqlEx)
                                {
                                    Error(sqlEx, "SubiektService", "Błąd podczas wyszukiwania przez SQL po błędzie NIP");
                                }
                            }
                        }
                        else
                        {
                            Info("Nie podano NIP - próba wyszukania kontrahenta przez SQL...", "SubiektService");
                            
                            // Próbuj wyszukać kontrahenta przez SQL nawet gdy nie ma NIP
                            try
                            {
                                int? kontrahentId = WyszukajKontrahentaPrzezSQL(null, email, customerName, phone, company, address, address1, address2, postcode, city, country, isoCode2);
                                
                                // Sprawdź czy użytkownik anulował (wartość -1 oznacza anulowanie)
                                if (kontrahentId.HasValue && kontrahentId.Value == -1)
                                {
                                    Info("Użytkownik anulował wybór kontrahenta - przerywam otwieranie ZK.", "SubiektService");
                                    return; // Przerwij otwieranie ZK
                                }
                                
                                if (kontrahentId.HasValue && kontrahentId.Value > 0 && zkDokument != null)
                                {
                                    // Ustaw ID kontrahenta w dokumencie
                                    int id = kontrahentId!.Value; // HasValue jest true, więc Value nie jest null
                                    zkDokument!.KontrahentId = id; // Sprawdziliśmy != null wyżej
                                    Info("Ustawiono kontrahenta o ID={id} (znaleziony przez SQL bez NIP) w dokumencie ZK.", "SubiektService");
                                }
                                else
                                {
                                    Info("Nie znaleziono kontrahenta przez SQL - dokument ZK zostanie otwarty bez kontrahenta.", "SubiektService");
                                }
                            }
                            catch (Exception sqlEx)
                            {
                                Error(sqlEx, "SubiektService", "Błąd podczas wyszukiwania przez SQL (bez NIP)");
                                Info("Dokument ZK zostanie otwarty bez kontrahenta.", "SubiektService");
                            }
                        }
                        
                        // Ustaw numer zamówienia jako numer oryginalnego dokumentu
                        if (!string.IsNullOrWhiteSpace(orderId) && zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.NumerOryginalny = orderId;
                                Info("Ustawiono numer oryginalnego dokumentu: {orderId}", "SubiektService");
                            }
                            catch (Exception numOrigEx)
                            {
                                Warning($"Nie udało się ustawić NumerOryginalny: {numOrigEx.Message}", "SubiektService");
                            }
                        }
                        
                        // Ustaw walutę dokumentu (jeśli została podana)
                        // Używamy atrybutu WalutaSymbol, który przyjmuje symbol waluty jako string (np. "EUR", "USD", "PLN")
                        if (!string.IsNullOrWhiteSpace(currency) && zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.WalutaSymbol = currency;
                                Info("Ustawiono walutę dokumentu ZK: {currency}", "SubiektService");
                                
                                // Opcjonalnie można pobrać kurs waluty automatycznie po ustawieniu symbolu
                                try
                                {
                                    zkDokument!.PobierzKursWaluty();
                                    Info("Pobrano kurs waluty dla {currency}", "SubiektService");
                                }
                                catch (Exception kursEx)
                                {
                                    Warning($"Nie udało się pobrać kursu waluty automatycznie: {kursEx.Message}", "SubiektService");
                                    // To nie jest błąd krytyczny, kurs można ustawić ręcznie później
                                }
                            }
                            catch (Exception walutaEx)
                            {
                                Warning($"Nie udało się ustawić waluty: {walutaEx.Message}", "SubiektService");
                            }
                        }
                        
                        // Ustaw przeliczanie dokumentu od cen brutto (ponieważ ustawiamy CenaBruttoPoRabacie dla pozycji)
                        if (zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.LiczonyOdCenBrutto = true;
                                Info("Ustawiono przeliczanie dokumentu ZK od cen brutto (LiczoneOdCenBrutto = true)", "SubiektService");
                            }
                            catch (Exception bruttoEx)
                            {
                                Warning($"Nie udało się ustawić LiczoneOdCenBrutto: {bruttoEx.Message}", "SubiektService");
                            }
                        }
                        
                        // Dodaj informacje o kuponie i numerze zamówienia do uwag dokumentu (jeśli jest kupon)
                        if ((!string.IsNullOrWhiteSpace(couponTitle) && couponAmount.HasValue) || !string.IsNullOrWhiteSpace(orderId))
                        {
                            try
                            {
                                var notes = new List<string>();
                                
                                // Dodaj numer zamówienia
                                if (!string.IsNullOrWhiteSpace(orderId))
                                {
                                    notes.Add($"Zamówienie: {orderId}");
                                }
                                
                                // Dodaj kupon
                                if (!string.IsNullOrWhiteSpace(couponTitle) && couponAmount.HasValue)
                                {
                                    string walutaNotatki = !string.IsNullOrWhiteSpace(currency) ? currency : "PLN";
                                    notes.Add($"{couponTitle} ({couponAmount.Value:F2} {walutaNotatki})");
                                }
                                
                                var fullNote = string.Join(" | ", notes);
                                
                                // Spróbuj użyć metody ZmienOpisDokumentu z SuDokumentyManager
                                try
                                {
                                    dynamic dokumentyManager = subiekt.SuDokumentyManager;
                                    if (dokumentyManager != null)
                                    {
                                        // sdoUwagi = 1 (wartość enum SuDokumentOpisEnum)
                                        const int sdoUwagi = 1;
                                        dokumentyManager.ZmienOpisDokumentu(zkDokument, sdoUwagi, fullNote);
                                        Info("Dodano informacje do uwag ZK (ZmienOpisDokumentu): {fullNote}", "SubiektService");
                                    }
                                }
                                catch (Exception methodEx)
                                {
                                    Warning($"ZmienOpisDokumentu nie zadziałała: {methodEx.Message}", "SubiektService");
                                    
                                    // Fallback: spróbuj bezpośrednio ustawić Uwagi
                                    if (zkDokument != null)
                                    {
                                        try
                                        {
                                            zkDokument!.Uwagi = fullNote;
                                            Info("Dodano informacje do uwag ZK (Uwagi): {fullNote}", "SubiektService");
                                        }
                                        catch (Exception directEx)
                                        {
                                            Warning($"Uwagi nie zadziałały: {directEx.Message}", "SubiektService");
                                            
                                            // Ostatni fallback: Opis
                                            try
                                            {
                                                zkDokument!.Opis = fullNote;
                                                Info("Dodano informacje do uwag ZK (Opis): {fullNote}", "SubiektService");
                                            }
                                            catch
                                            {
                                                Info("Nie udało się ustawić żadnego pola uwag dla dokumentu ZK.", "SubiektService");
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception couponEx)
                            {
                                Warning($"Błąd dodawania informacji do uwag ZK: {couponEx.Message}", "SubiektService");
                            }
                        }
                        // Najpierw oblicz procent kuponu (jeśli jest kupon), aby zastosować go później na pozycjach
                        double? couponPercentage = null;
                        if (couponAmount.HasValue && items != null)
                        {
                            double totalProductsValue = 0.0;
                            foreach (var it in items)
                            {
                                if (!string.IsNullOrWhiteSpace(it.Price))
                                {
                                    if (double.TryParse(it.Price, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var price))
                                    {
                                        totalProductsValue += price * it.Quantity;
                                    }
                                }
                            }
                            
                            if (totalProductsValue > 0.0)
                            {
                                couponPercentage = (couponAmount.Value / totalProductsValue) * 100.0;
                                Info("ANALIZA KUPONU:", "SubiektService");
                                Info("Wartość kuponu: {couponAmount.Value:F2}", "SubiektService");
                                Info("Suma wartości produktów: {totalProductsValue:F2}", "SubiektService");
                                Info("Procent kuponu względem produktów: {couponPercentage:F2}%", "SubiektService");
                            }
                            else
                            {
                                Info("Brak danych o cenach produktów - nie można obliczyć procentu kuponu.", "SubiektService");
                            }
                        }
                        
                        // Dodaj pozycje z listy produktów (product_id)
                        if (items != null)
                        {
                            try
                            {
                                if (zkDokument != null)
                                {
                                    dynamic pozycje = zkDokument!.Pozycje;
                                    foreach (var it in items)
                                    {
                                    if (int.TryParse(it.ProductId, out var towarId))
                                    {
                                        dynamic pozycja = pozycje.Dodaj(towarId);

                                        // Ilość
                                        if (it.Quantity > 0)
                                        {
                                            try { pozycja.IloscJm = it.Quantity; } catch { }
                                        }

                                        // Oblicz cenę brutto z API (Price + Tax)
                                        double? apiPriceBrutto = null;
                                        if (!string.IsNullOrWhiteSpace(it.Price))
                                        {
                                            try
                                            {
                                                var priceNetto = double.Parse(it.Price, System.Globalization.CultureInfo.InvariantCulture);
                                                var tax = 0.0;
                                                
                                                // Dodaj podatek jeśli jest dostępny
                                                if (!string.IsNullOrWhiteSpace(it.Tax))
                                                {
                                                    if (double.TryParse(it.Tax, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedTax))
                                                    {
                                                        tax = parsedTax;
                                                    }
                                                }
                                                
                                                // Oblicz cenę brutto
                                                var priceBrutto = priceNetto + tax;
                                                
                                                // Odejmij procent kuponu od ceny brutto (jeśli kupon istnieje)
                                                if (couponPercentage.HasValue)
                                                {
                                                    priceBrutto = priceBrutto * (1.0 - couponPercentage.Value / 100.0);
                                                }
                                                
                                                apiPriceBrutto = priceBrutto;
                                            }
                                            catch (Exception priceEx)
                                            {
                                                Error(priceEx, "SubiektService", $"Błąd parsowania ceny dla produktu {it.ProductId}");
                                                apiPriceBrutto = null;
                                            }
                                        }

                                        // Ustaw cenę brutto po rabacie z API
                                        if (apiPriceBrutto.HasValue)
                                        {
                                            try
                                            {
                                                pozycja.CenaBruttoPoRabacie = apiPriceBrutto.Value;
                                                Info("Ustawiono CenaBruttoPoRabacie={apiPriceBrutto.Value:F2} dla produktu ID={towarId} (Price={it.Price}, Tax={it.Tax})", "SubiektService");
                                            }
                                            catch (Exception cenaEx)
                                            {
                                                Warning($"Nie udało się ustawić CenaBruttoPoRabacie: {cenaEx.Message}", "SubiektService");
                                            }
                                        }
                                        else
                                        {
                                            Info("Brak ceny brutto z API dla produktu ID={towarId}", "SubiektService");
                                        }
                                        
                                        Debug($"Dodano pozycję towarową o ID={towarId} (qty={it.Quantity}, apiPriceNetto={it.Price}, apiPriceBrutto={apiPriceBrutto?.ToString("F2") ?? "brak"}).", "SubiektService");
                                    }
                                    else
                                    {
                                        Info("Pominięto pozycję z nieprawidłowym product_id: '{it.ProductId}'", "SubiektService");
                                    }
                                    }
                                }
                            }
                            catch (Exception dodajPozEx)
                            {
                                Error(dodajPozEx, "SubiektService", "Nie udało się dodać pozycji");
                            }
                        }
                        
                        

                        // Dodaj towar 'KOSZTY/1' na końcu
                        // Suma handling + gls (oba używają KOSZTY/1)
                        double sumaKosztowKOSZTY1 = 0.0;
                        if (handlingAmount.HasValue)
                        {
                            sumaKosztowKOSZTY1 += handlingAmount.Value;
                        }
                        if (glsAmount.HasValue)
                        {
                            sumaKosztowKOSZTY1 += glsAmount.Value;
                        }
                        
                        if (sumaKosztowKOSZTY1 > 0.0)
                        {
                            try
                            {
                                dynamic towary = subiekt.Towary;
                                dynamic towarKoszty = towary.Wczytaj("KOSZTY/1");
                                
                                if (towarKoszty != null && zkDokument != null)
                                {
                                    dynamic pozycje = zkDokument!.Pozycje;
                                    dynamic kosztyPoz = pozycje.Dodaj(towarKoszty!.Identyfikator);
                                    try { kosztyPoz.IloscJm = 1; } catch { }
                                    // Ustaw cenę na sumę handling + gls
                                    try 
                                    { 
                                        kosztyPoz.CenaNettoPrzedRabatem = sumaKosztowKOSZTY1;
                                        var skladniki = new System.Collections.Generic.List<string>();
                                        if (handlingAmount.HasValue && handlingAmount.Value > 0) skladniki.Add($"handling ({handlingAmount.Value:F2})");
                                        if (glsAmount.HasValue && glsAmount.Value > 0) skladniki.Add($"gls ({glsAmount.Value:F2})");
                                        Info($"Dodano towar 'KOSZTY/1' z ceną {sumaKosztowKOSZTY1:F2} (składniki: {string.Join(" + ", skladniki)}).", "SubiektService");
                                    }
                                    catch
                                    {
                                        Info("Dodano towar 'KOSZTY/1' na końcu pozycji, ale nie udało się ustawić ceny.", "SubiektService");
                                    }
                                }
                                else
                                {
                                    Info("Nie znaleziono towaru o symbolu 'KOSZTY/1'.", "SubiektService");
                                }
                            }
                            catch (Exception kosztyEx)
                            {
                                Warning($"Nie udało się dodać towaru 'KOSZTY/1': {kosztyEx.Message}", "SubiektService");
                            }
                        }
                        else
                        {
                            Info("Brak handling i gls w API - pomijam dodawanie 'KOSZTY/1'.", "SubiektService");
                        }
                        
                        // Dodaj towar 'KOSZTY/2' na końcu
                        // Tylko jeśli w API jest shipping
                        if (shippingAmount.HasValue)
                        {
                            try
                            {
                                dynamic towary = subiekt.Towary;
                                dynamic towarShipping = towary.Wczytaj("KOSZTY/2");
                                
                                if (towarShipping != null && zkDokument != null)
                                {
                                    dynamic pozycje = zkDokument!.Pozycje;
                                    dynamic shippingPoz = pozycje.Dodaj(towarShipping!.Identyfikator);
                                    try { shippingPoz.IloscJm = 1; } catch { }
                                    // Ustaw cenę na wartość shipping z API
                                    try 
                                    { 
                                        shippingPoz.CenaNettoPrzedRabatem = shippingAmount.Value;
                                        Info("Dodano towar 'KOSZTY/2' z ceną shipping ({shippingAmount.Value:F2}).", "SubiektService");
                                    }
                                    catch
                                    {
                                        Info("Dodano towar 'KOSZTY/2' na końcu pozycji, ale nie udało się ustawić ceny.", "SubiektService");
                                    }
                                }
                                else
                                {
                                    Info("Nie znaleziono towaru o symbolu 'KOSZTY/2'.", "SubiektService");
                                }
                            }
                            catch (Exception shippingEx)
                            {
                                Warning($"Nie udało się dodać towaru 'KOSZTY/2': {shippingEx.Message}", "SubiektService");
                            }
                        }
                        else
                        {
                            Info("Brak shipping w API - pomijam dodawanie 'KOSZTY/2'.", "SubiektService");
                        }
                        
                        // Dodaj towar 'KOSZTY/2' na końcu
                        // Tylko jeśli w API jest cod_fee
                        if (codFeeAmount.HasValue)
                        {
                            try
                            {
                                dynamic towary = subiekt.Towary;
                                dynamic towarCodFee = towary.Wczytaj("KOSZTY/2");
                                
                                if (towarCodFee != null && zkDokument != null)
                                {
                                    dynamic pozycje = zkDokument!.Pozycje;
                                    dynamic codFeePoz = pozycje.Dodaj(towarCodFee!.Identyfikator);
                                    try { codFeePoz.IloscJm = 1; } catch { }
                                    // Ustaw cenę na wartość cod_fee z API
                                    try 
                                    { 
                                        codFeePoz.CenaNettoPrzedRabatem = codFeeAmount.Value;
                                        Info("Dodano towar 'KOSZTY/2' z ceną cod_fee ({codFeeAmount.Value:F2}).", "SubiektService");
                                    }
                                    catch
                                    {
                                        Info("Dodano towar 'KOSZTY/2' na końcu pozycji, ale nie udało się ustawić ceny.", "SubiektService");
                                    }
                                }
                                else
                                {
                                    Info("Nie znaleziono towaru o symbolu 'KOSZTY/2'.", "SubiektService");
                                }
                            }
                            catch (Exception codFeeEx)
                            {
                                Warning($"Nie udało się dodać towaru 'KOSZTY/2' dla cod_fee: {codFeeEx.Message}", "SubiektService");
                            }
                        }
                        else
                        {
                            Info("Brak cod_fee w API - pomijam dodawanie 'KOSZTY/2' dla cod_fee.", "SubiektService");
                        }
                        
                        // Przelicz dokument przed odczytem wartości (może być wymagane)
                        if (zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.Przelicz();
                                Info("Dokument ZK został przeliczony przed odczytem wartości pozycji.", "SubiektService");
                            }
                            catch (Exception przeliczEx)
                            {
                                Warning($"Nie udało się przeliczyć dokumentu przed odczytem wartości: {przeliczEx.Message}", "SubiektService");
                            }
                        }
                        
                        // Odczytaj sumę wartości wszystkich pozycji z dokumentu ZK
                        if (zkDokument != null)
                        {
                            try
                            {
                                double sumaWartosci = 0.0;
                                dynamic pozycje = zkDokument.Pozycje;
                            
                            // Przejdź przez wszystkie pozycje i zsumuj ich wartości
                            // W Subiekcie GT indeksy pozycji zaczynają się od 1, nie od 0!
                            int liczbaPozycji = pozycje.Liczba;
                            Info("Liczba pozycji w dokumencie ZK: {liczbaPozycji}", "SubiektService");
                            
                            for (int i = 1; i <= liczbaPozycji; i++)
                            {
                                try
                                {
                                    dynamic pozycja = pozycje.Element(i);
                                    
                                    // Odczytaj wartość pozycji bezpośrednio z WartoscBruttoPoRabacie
                                    double wartoscPozycji = 0.0;
                                    try
                                    {
                                        var wartosc = pozycja.WartoscBruttoPoRabacie;
                                        wartoscPozycji = Convert.ToDouble(wartosc);
                                    }
                                    catch (Exception wartoscEx)
                                    {
                                        Warning($"Błąd odczytu WartoscBruttoPoRabacie dla pozycji {i}: {wartoscEx.Message}", "SubiektService");
                                    }
                                    
                                    sumaWartosci += wartoscPozycji;
                                    Info("Pozycja {i}: Wartość = {wartoscPozycji:F2}", "SubiektService");
                                    
                                    // Dodatkowe logowanie - sprawdź jakie właściwości pozycji są dostępne
                                    try
                                    {
                                        var ilosc = pozycja.IloscJm;
                                        Info("  - Ilość: {ilosc}", "SubiektService");
                                    }
                                    catch { }
                                    try
                                    {
                                        var cenaBruttoPoRabacie = pozycja.CenaBruttoPoRabacie;
                                        Info("  - CenaBruttoPoRabacie: {cenaBruttoPoRabacie:F2}", "SubiektService");
                                    }
                                    catch { }
                                    try
                                    {
                                        var cenaNettoPoRabacie = pozycja.CenaNettoPoRabacie;
                                        Info("  - CenaNettoPoRabacie: {cenaNettoPoRabacie:F2}", "SubiektService");
                                    }
                                    catch { }
                                }
                                catch (Exception pozEx)
                                {
                                    Warning($"Błąd odczytu pozycji {i}: {pozEx.Message}", "SubiektService");
                                }
                            }
                            
                            Info("=========================================", "SubiektService");
                            Info("SUMA WARTOŚCI WSZYSTKICH POZYCJI: {sumaWartosci:F2}", "SubiektService");
                            Info("=========================================", "SubiektService");
                            
                            // Porównaj sumę pozycji ZK z wartością zamówienia z API
                            if (!string.IsNullOrWhiteSpace(orderTotal))
                            {
                                try
                                {
                                    if (double.TryParse(orderTotal, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double orderTotalValue))
                                    {
                                        double roznica = Math.Abs(sumaWartosci - orderTotalValue);
                                        double roznicaProcent = sumaWartosci > 0 ? (roznica / sumaWartosci) * 100.0 : 0.0;
                                        
                                        Info("Wartość zamówienia z API: {orderTotalValue:F2}", "SubiektService");
                                        Info("Różnica: {roznica:F2} ({roznicaProcent:F2}%)", "SubiektService");
                                        
                                        // Jeśli różnica jest większa niż próg tolerancji (0.001 = 0.01 grosza)
                                        // Używamy progu zamiast > 0.0 aby uniknąć błędów zaokrąglenia zmiennoprzecinkowych
                                        if (roznica > 0.001)
                                        {
                                            // Analiza przyczyn różnicy
                                            Info("=========================================", "SubiektService");
                                            Info("ANALIZA RÓŻNICY W WARTOŚCIACH", "SubiektService");
                                            Info("=========================================", "SubiektService");
                                            
                                            var wnioski = new System.Collections.Generic.List<string>();
                                            
                                            // Porównaj bezpośrednio sumę ZK z wartością z API
                                            // Wartość z API jest już finalną wartością (uwzględnia wszystkie składniki)
                                            Info("Suma pozycji ZK: {sumaWartosci:F2}", "SubiektService");
                                            Info("Wartość zamówienia z API (total): {orderTotalValue:F2}", "SubiektService");
                                            Info("Różnica: {roznica:F2} ({roznicaProcent:F2}%)", "SubiektService");
                                            
                                            // Sprawdź komponenty zamówienia z API (dla informacji)
                                            double sumaKosztow = 0.0;
                                            if (handlingAmount.HasValue && handlingAmount.Value > 0.01)
                                            {
                                                sumaKosztow += handlingAmount.Value;
                                                Info("Handling z API: {handlingAmount.Value:F2}", "SubiektService");
                                            }
                                            if (glsAmount.HasValue && glsAmount.Value > 0.01)
                                            {
                                                sumaKosztow += glsAmount.Value;
                                                Info("Gls z API: {glsAmount.Value:F2}", "SubiektService");
                                            }
                                            if (shippingAmount.HasValue && shippingAmount.Value > 0.01)
                                            {
                                                sumaKosztow += shippingAmount.Value;
                                                Info("Shipping z API: {shippingAmount.Value:F2}", "SubiektService");
                                            }
                                            if (codFeeAmount.HasValue && codFeeAmount.Value > 0.01)
                                            {
                                                sumaKosztow += codFeeAmount.Value;
                                                Info("Cod_fee z API: {codFeeAmount.Value:F2}", "SubiektService");
                                            }
                                            if (sumaKosztow > 0.01)
                                            {
                                                Info("Suma kosztów z API: {sumaKosztow:F2}", "SubiektService");
                                            }
                                            
                                            if (couponAmount.HasValue && couponAmount.Value > 0.01)
                                            {
                                                Info("Kupon z API: -{couponAmount.Value:F2}", "SubiektService");
                                            }
                                            
                                            // Analiza różnicy między ZK a API
                                            if (roznica > 0.01)
                                            {
                                                // Sprawdź czy ZK ma mniej niż API (może brakować pozycji)
                                                if (sumaWartosci < orderTotalValue)
                                                {
                                                    double brakujacaKwota = orderTotalValue - sumaWartosci;
                                                    wnioski.Add($"Suma ZK jest mniejsza o {brakujacaKwota:F2} - prawdopodobnie brakuje pozycji kosztów.");
                                                    
                                                    // Sprawdź które koszty mogą brakować
                                                    if ((handlingAmount.HasValue && handlingAmount.Value > 0.01) || (glsAmount.HasValue && glsAmount.Value > 0.01))
                                                    {
                                                        double sumaKosztowKOSZTY1Wnioski = (handlingAmount ?? 0) + (glsAmount ?? 0);
                                                        wnioski.Add($"  Sprawdź czy dodano pozycję 'KOSZTY/1' (handling: {handlingAmount ?? 0:F2} + gls: {glsAmount ?? 0:F2} = {sumaKosztowKOSZTY1Wnioski:F2}).");
                                                    }
                                                    if (shippingAmount.HasValue && shippingAmount.Value > 0.01)
                                                    {
                                                        wnioski.Add($"  Sprawdź czy dodano pozycję 'KOSZTY/2' dla shipping ({shippingAmount.Value:F2}).");
                                                    }
                                                    if (codFeeAmount.HasValue && codFeeAmount.Value > 0.01)
                                                    {
                                                        wnioski.Add($"  Sprawdź czy dodano pozycję 'KOSZTY/2' dla cod_fee ({codFeeAmount.Value:F2}).");
                                                    }
                                                }
                                                // Sprawdź czy ZK ma więcej niż API (może być różnica w cenach lub dodatkowe pozycje)
                                                else if (sumaWartosci > orderTotalValue)
                                                {
                                                    double nadmiar = sumaWartosci - orderTotalValue;
                                                    wnioski.Add($"Suma ZK jest większa o {nadmiar:F2} - możliwe różnice w cenach produktów, VAT lub dodatkowe pozycje.");
                                                    
                                                    // Sprawdź liczbę pozycji
                                                    if (items != null && items.Any())
                                                    {
                                                        int liczbaPozycjiZK = pozycje.Liczba;
                                                        int liczbaProduktow = items.Count();
                                                        int liczbaKosztow = 0;
                                                        if ((handlingAmount.HasValue && handlingAmount.Value > 0.01) || (glsAmount.HasValue && glsAmount.Value > 0.01)) liczbaKosztow++; // KOSZTY/1 dla handling+gls
                                                        if (shippingAmount.HasValue && shippingAmount.Value > 0.01) liczbaKosztow++; // KOSZTY/2 dla shipping
                                                        if (codFeeAmount.HasValue && codFeeAmount.Value > 0.01) liczbaKosztow++; // KOSZTY/2 dla cod_fee
                                                        
                                                        int oczekiwanaLiczbaPozycji = liczbaProduktow + liczbaKosztow;
                                                        if (liczbaPozycjiZK != oczekiwanaLiczbaPozycji)
                                                        {
                                                            wnioski.Add($"  Liczba pozycji w ZK ({liczbaPozycjiZK}) różni się od oczekiwanej ({oczekiwanaLiczbaPozycji} - {liczbaProduktow} produktów + {liczbaKosztow} kosztów).");
                                                        }
                                                    }
                                                    
                                                    wnioski.Add($"  Możliwe przyczyny: różnice w zaokrągleniach VAT, różne ceny w Subiekcie vs API, lub dodatkowe pozycje w ZK.");
                                                }
                                            }
                                            else
                                            {
                                                wnioski.Add("Wartości się zgadzają (różnica = 0.00).");
                                            }
                                            
                                            // Wyświetl wnioski
                                            Info("=========================================", "SubiektService");
                                            Info("WNIOSKI:", "SubiektService");
                                            foreach (var wniosek in wnioski)
                                            {
                                                Info("  - {wniosek}", "SubiektService");
                                            }
                                            Info("=========================================", "SubiektService");
                                            
                                            // Użyj waluty z API, jeśli jest dostępna, w przeciwnym razie domyślnie PLN
                                            string walutaDisplay = !string.IsNullOrWhiteSpace(currency) ? currency : "PLN";
                                            
                                            string komunikat = $"Suma wartości pozycji w dokumencie ZK: {sumaWartosci:F2} {walutaDisplay}\nWartość zamówienia z API: {orderTotalValue:F2} {walutaDisplay}\n\nRóżnica: {roznica:F2} {walutaDisplay} ({roznicaProcent:F2}%)\n\nKorekta zostanie wykonana w następującej kolejności:\n1. Pozycja KOSZTY/1 (jeśli istnieje)\n2. Pozycja KOSZTY/2 (jeśli istnieje)\n3. Pierwsza pozycja produktowa";
                                            
                                            var dialog = new Gryzak.Views.KorektaWartosciDialog(komunikat);
                                            bool czyKorygowac = dialog.ShowDialog() == true && dialog.CzyKorygowac;
                                            
                                            // Sprawdź czy użytkownik anulował - jeśli tak, przerwij otwieranie ZK
                                            if (dialog.CzyAnulowac)
                                            {
                                                Info("Użytkownik anulował w oknie korekty - przerywam otwieranie ZK.", "SubiektService");
                                                return; // Przerwij otwieranie ZK
                                            }
                                            
                                            if (czyKorygowac)
                                            {
                                                // Koryguj różnicę w kolejności: KOSZTY/1 -> KOSZTY/2 -> pierwsza pozycja produktowa
                                                try
                                                {
                                                    if (pozycje.Liczba > 0)
                                                    {
                                                        // Pobierz identyfikatory towarów KOSZTY/1 i KOSZTY/2
                                                        int? towarKOSZTY1Id = null;
                                                        int? towarKOSZTY2Id = null;
                                                        
                                                        try
                                                        {
                                                            dynamic towary = subiekt.Towary;
                                                            dynamic towarKOSZTY1 = towary.Wczytaj("KOSZTY/1");
                                                            if (towarKOSZTY1 != null)
                                                            {
                                                                towarKOSZTY1Id = towarKOSZTY1.Identyfikator;
                                                            }
                                                            dynamic towarKOSZTY2 = towary.Wczytaj("KOSZTY/2");
                                                            if (towarKOSZTY2 != null)
                                                            {
                                                                towarKOSZTY2Id = towarKOSZTY2.Identyfikator;
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            Info("Nie udało się odczytać identyfikatorów towarów KOSZTY.", "SubiektService");
                                                        }
                                                        
                                                        double pozostalaRoznica = roznica;
                                                        bool czyZKwieksze = sumaWartosci > orderTotalValue;
                                                        
                                                        // Krok 1: Spróbuj skorygować KOSZTY/1 jeśli istnieje
                                                        if (towarKOSZTY1Id.HasValue && pozostalaRoznica > 0.001)
                                                        {
                                                            for (int i = 1; i <= pozycje.Liczba; i++)
                                                            {
                                                                try
                                                                {
                                                                    dynamic poz = pozycje.Element(i);
                                                                    int? towarId = null;
                                                                    try
                                                                    {
                                                                        dynamic towar = poz.Towar;
                                                                        if (towar != null)
                                                                        {
                                                                            towarId = towar.Identyfikator;
                                                                        }
                                                                    }
                                                                    catch
                                                                    {
                                                                        try
                                                                        {
                                                                            towarId = poz.IdentyfikatorTowaru;
                                                                        }
                                                                        catch { }
                                                                    }
                                                                    
                                                                    if (towarId.HasValue && towarId.Value == towarKOSZTY1Id.Value)
                                                                    {
                                                                        // To jest pozycja KOSZTY/1 - skoryguj ją
                                                                        double aktualnaWartosc = 0.0;
                                                                        try
                                                                        {
                                                                            var wartosc = poz.WartoscBruttoPoRabacie;
                                                                            aktualnaWartosc = Convert.ToDouble(wartosc);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        double ilosc = 1.0;
                                                                        try
                                                                        {
                                                                            ilosc = Convert.ToDouble(poz.IloscJm);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        if (ilosc > 0)
                                                                        {
                                                                            double nowaWartosc = czyZKwieksze
                                                                                ? Math.Max(0, aktualnaWartosc - pozostalaRoznica)
                                                                                : aktualnaWartosc + pozostalaRoznica;
                                                                            double nowaCenaBrutto = nowaWartosc / ilosc;
                                                                            
                                                                            poz.CenaBruttoPoRabacie = nowaCenaBrutto;
                                                                            zkDokument.Przelicz();
                                                                            
                                                                            Info("Skorygowano pozycję KOSZTY/1 (pozycja {i}):", "SubiektService");
                                                                            Info("  Stara wartość: {aktualnaWartosc:F2}", "SubiektService");
                                                                            Debug($"  Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}", "SubiektService");
                                                                            Info("  Nowa wartość: {nowaWartosc:F2}", "SubiektService");
                                                                            
                                                                            // Oblicz ile różnicy zostało skorygowane
                                                                            double skorygowanaRoznica = Math.Abs(aktualnaWartosc - nowaWartosc);
                                                                            pozostalaRoznica -= skorygowanaRoznica;
                                                                            
                                                                            if (pozostalaRoznica <= 0.001) break;
                                                                        }
                                                                    }
                                                                }
                                                                catch { }
                                                            }
                                                        }
                                                        
                                                        // Krok 2: Spróbuj skorygować KOSZTY/2 jeśli istnieje i różnica nadal jest
                                                        if (towarKOSZTY2Id.HasValue && pozostalaRoznica > 0.001)
                                                        {
                                                            for (int i = 1; i <= pozycje.Liczba; i++)
                                                            {
                                                                try
                                                                {
                                                                    dynamic poz = pozycje.Element(i);
                                                                    int? towarId = null;
                                                                    try
                                                                    {
                                                                        dynamic towar = poz.Towar;
                                                                        if (towar != null)
                                                                        {
                                                                            towarId = towar.Identyfikator;
                                                                        }
                                                                    }
                                                                    catch
                                                                    {
                                                                        try
                                                                        {
                                                                            towarId = poz.IdentyfikatorTowaru;
                                                                        }
                                                                        catch { }
                                                                    }
                                                                    
                                                                    if (towarId.HasValue && towarId.Value == towarKOSZTY2Id.Value)
                                                                    {
                                                                        // To jest pozycja KOSZTY/2 - skoryguj ją
                                                                        double aktualnaWartosc = 0.0;
                                                                        try
                                                                        {
                                                                            var wartosc = poz.WartoscBruttoPoRabacie;
                                                                            aktualnaWartosc = Convert.ToDouble(wartosc);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        double ilosc = 1.0;
                                                                        try
                                                                        {
                                                                            ilosc = Convert.ToDouble(poz.IloscJm);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        if (ilosc > 0)
                                                                        {
                                                                            double nowaWartosc = czyZKwieksze
                                                                                ? Math.Max(0, aktualnaWartosc - pozostalaRoznica)
                                                                                : aktualnaWartosc + pozostalaRoznica;
                                                                            double nowaCenaBrutto = nowaWartosc / ilosc;
                                                                            
                                                                            poz.CenaBruttoPoRabacie = nowaCenaBrutto;
                                                                            zkDokument.Przelicz();
                                                                            
                                                                            Info("Skorygowano pozycję KOSZTY/2 (pozycja {i}):", "SubiektService");
                                                                            Info("  Stara wartość: {aktualnaWartosc:F2}", "SubiektService");
                                                                            Debug($"  Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}", "SubiektService");
                                                                            Info("  Nowa wartość: {nowaWartosc:F2}", "SubiektService");
                                                                            
                                                                            // Oblicz ile różnicy zostało skorygowane
                                                                            double skorygowanaRoznica = Math.Abs(aktualnaWartosc - nowaWartosc);
                                                                            pozostalaRoznica -= skorygowanaRoznica;
                                                                            
                                                                            if (pozostalaRoznica <= 0.001) break;
                                                                        }
                                                                    }
                                                                }
                                                                catch { }
                                                            }
                                                        }
                                                        
                                                        // Krok 3: Jeśli nadal jest różnica, skoryguj pierwszą pozycję produktową (pierwsza nie-KOSZTY)
                                                        if (pozostalaRoznica > 0.001)
                                                        {
                                                            for (int i = 1; i <= pozycje.Liczba; i++)
                                                            {
                                                                try
                                                                {
                                                                    dynamic poz = pozycje.Element(i);
                                                                    int? towarId = null;
                                                                    bool jestKoszty = false;
                                                                    try
                                                                    {
                                                                        dynamic towar = poz.Towar;
                                                                        if (towar != null)
                                                                        {
                                                                            towarId = towar.Identyfikator;
                                                                            if (towarId.HasValue && towarKOSZTY1Id.HasValue && towarId.Value == towarKOSZTY1Id.Value)
                                                                            {
                                                                                jestKoszty = true;
                                                                            }
                                                                            if (towarId.HasValue && towarKOSZTY2Id.HasValue && towarId.Value == towarKOSZTY2Id.Value)
                                                                            {
                                                                                jestKoszty = true;
                                                                            }
                                                                        }
                                                                    }
                                                                    catch
                                                                    {
                                                                        try
                                                                        {
                                                                            towarId = poz.IdentyfikatorTowaru;
                                                                            if (towarId.HasValue && towarKOSZTY1Id.HasValue && towarId.Value == towarKOSZTY1Id.Value)
                                                                            {
                                                                                jestKoszty = true;
                                                                            }
                                                                            if (towarId.HasValue && towarKOSZTY2Id.HasValue && towarId.Value == towarKOSZTY2Id.Value)
                                                                            {
                                                                                jestKoszty = true;
                                                                            }
                                                                        }
                                                                        catch { }
                                                                    }
                                                                    
                                                                    if (!jestKoszty)
                                                                    {
                                                                        // To jest pozycja produktowa - skoryguj ją
                                                                        double aktualnaWartosc = 0.0;
                                                                        try
                                                                        {
                                                                            var wartosc = poz.WartoscBruttoPoRabacie;
                                                                            aktualnaWartosc = Convert.ToDouble(wartosc);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        double ilosc = 0.0;
                                                                        try
                                                                        {
                                                                            ilosc = Convert.ToDouble(poz.IloscJm);
                                                                        }
                                                                        catch { }
                                                                        
                                                                        if (ilosc > 0)
                                                                        {
                                                                            double nowaWartosc = czyZKwieksze
                                                                                ? aktualnaWartosc - pozostalaRoznica
                                                                                : aktualnaWartosc + pozostalaRoznica;
                                                                            double nowaCenaBrutto = nowaWartosc / ilosc;
                                                                            
                                                                            poz.CenaBruttoPoRabacie = nowaCenaBrutto;
                                                                            zkDokument.Przelicz();
                                                                            
                                                                            Info("Skorygowano pierwszą pozycję produktową (pozycja {i}):", "SubiektService");
                                                                            Info("  Stara wartość: {aktualnaWartosc:F2}", "SubiektService");
                                                                            Debug($"  Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}", "SubiektService");
                                                                            Info("  Nowa wartość: {nowaWartosc:F2}", "SubiektService");
                                                                            Info("  Nowa cena brutto: {nowaCenaBrutto:F2}", "SubiektService");
                                                                            
                                                                            pozostalaRoznica = 0.0;
                                                                            break; // Korekta zakończona
                                                                        }
                                                                    }
                                                                }
                                                                catch { }
                                                            }
                                                        }
                                                        
                                                        // Przelicz dokument po wszystkich zmianach
                                                        zkDokument.Przelicz();
                                                        
                                                        // Odczytaj sumę ponownie po korekcie
                                                        double sumaPoKorekcie = 0.0;
                                                        for (int j = 1; j <= pozycje.Liczba; j++)
                                                        {
                                                            try
                                                            {
                                                                dynamic poz = pozycje.Element(j);
                                                                var wart = poz.WartoscBruttoPoRabacie;
                                                                sumaPoKorekcie += Convert.ToDouble(wart);
                                                            }
                                                            catch { }
                                                        }
                                                        
                                                        double roznicaPoKorekcie = Math.Abs(sumaPoKorekcie - orderTotalValue);
                                                        Info("Suma po korekcie: {sumaPoKorekcie:F2}", "SubiektService");
                                                        Info("Różnica po korekcie: {roznicaPoKorekcie:F2}", "SubiektService");
                                                        
                                                        if (roznicaPoKorekcie <= 0.001)
                                                        {
                                                            Info("Korekta zakończona pomyślnie. Wartości są teraz zgodne.", "SubiektService");
                                                        }
                                                        else
                                                        {
                                                            MessageBox.Show(
                                                                $"Korekta wykonana, ale różnica nadal wynosi {roznicaPoKorekcie:F2}.\n\nMożliwe przyczyny: zaokrąglenia lub różnice w VAT.",
                                                                "Błąd korekty",
                                                                System.Windows.MessageBoxButton.OK,
                                                                System.Windows.MessageBoxImage.Warning);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Info("Błąd: Brak pozycji w dokumencie - nie można wykonać korekty.", "SubiektService");
                                                        MessageBox.Show(
                                                            "Nie można skorygować - brak pozycji w dokumencie.",
                                                            "Błąd korekty",
                                                            System.Windows.MessageBoxButton.OK,
                                                            System.Windows.MessageBoxImage.Error);
                                                    }
                                                }
                                                catch (Exception cenaEx)
                                                {
                                                    Warning($"Błąd podczas korekty ceny: {cenaEx.Message}", "SubiektService");
                                                    MessageBox.Show(
                                                        $"Błąd podczas korekty: {cenaEx.Message}",
                                                        "Błąd korekty",
                                                        System.Windows.MessageBoxButton.OK,
                                                        System.Windows.MessageBoxImage.Error);
                                                }
                                            }
                                            else
                                            {
                                                Info("Użytkownik wybrał 'Pozostaw' - różnica nie została skorygowana.", "SubiektService");
                                            }
                                        }
                                        else
                                        {
                                            Info("Wartości się zgadzają (różnica = 0.00).", "SubiektService");
                                        }
                                    }
                                }
                                catch (Exception porownanieEx)
                                {
                                    Warning($"Błąd podczas porównywania wartości: {porownanieEx.Message}", "SubiektService");
                                }
                            }
                            else
                            {
                                Info("Brak wartości zamówienia z API do porównania.", "SubiektService");
                            }
                            
                            // Spróbuj też odczytać wartość bezpośrednio z dokumentu (jeśli dostępna)
                            if (zkDokument != null)
                            {
                                try
                                {
                                    double wartoscDokumentuBrutto = zkDokument.WartoscBrutto;
                                    Info("Wartość brutto dokumentu ZK (z właściwości): {wartoscDokumentuBrutto:F2}", "SubiektService");
                                }
                                catch { }
                            }
                            
                            if (zkDokument != null)
                            {
                                try
                                {
                                    double wartoscDokumentuNetto = zkDokument!.WartoscNetto;
                                    Info("Wartość netto dokumentu ZK (z właściwości): {wartoscDokumentuNetto:F2}", "SubiektService");
                                }
                                catch { }
                                
                                try
                                {
                                    double wartoscDokumentu = zkDokument!.Wartosc;
                                    Info("Wartość dokumentu ZK (z właściwości): {wartoscDokumentu:F2}", "SubiektService");
                                }
                                catch { }
                            }
                        }
                        catch (Exception sumaEx)
                        {
                            Warning($"Błąd podczas odczytu sumy wartości pozycji: {sumaEx.Message}", "SubiektService");
                        }

                        // Otwórz okno dokumentu używając metody Wyswietl() - zgodnie z przykładem
                        if (zkDokument != null)
                        {
                            try
                            {
                                // Wyswietl(false) - otwiera okno w trybie edycji
                                zkDokument!.Wyswietl(false);
                                Info("Wywołano Wyswietl(false) na dokumencie ZK.", "SubiektService");

                                // Aktywuj okno ZK na wierzch - używamy głównego okna Subiekta
                                try
                                {
                                    subiekt.Okno.Aktywuj();
                                    Info("Aktywowano główne okno Subiekta GT na wierzch.", "SubiektService");
                                }
                                catch (Exception aktywujEx)
                                {
                                    Warning($"Aktywacja głównego okna Subiekta nie zadziałała: {aktywujEx.Message}", "SubiektService");
                                }
                                
                                Info("Okno dokumentu ZK zostało otwarte pomyślnie.", "SubiektService");
                            }
                            catch (Exception wyswietlEx)
                            {
                                Error(wyswietlEx, "SubiektService", "BŁĄD: Wyswietl() nie zadziałał");
                                
                                // Spróbuj bez parametru
                                if (zkDokument != null)
                                {
                                    try
                                    {
                                        zkDokument.Wyswietl();
                                        Info("Wyswietl() bez parametru zadziałał.", "SubiektService");
                                        
                                        // Aktywuj główne okno Subiekta na wierzch
                                        try
                                        {
                                            subiekt.Okno.Aktywuj();
                                            Info("Aktywowano główne okno Subiekta na wierzch (fallback).", "SubiektService");
                                        }
                                        catch (Exception aktywujEx2)
                                        {
                                            Warning($"Aktywacja głównego okna Subiekta (fallback) nie zadziałała: {aktywujEx2.Message}", "SubiektService");
                                        }
                                    }
                                    catch (Exception wyswietl2Ex)
                                    {
                                        Error(wyswietl2Ex, "SubiektService", "Błąd: Wyswietl() bez parametru też nie zadziałał");
                                        MessageBox.Show(
                                            "Dokument ZK został utworzony w Subiekcie GT, ale nie udało się otworzyć okna edycji.\n\nSprawdź okno główne Subiekta GT.",
                                            "Informacja",
                                            System.Windows.MessageBoxButton.OK,
                                            System.Windows.MessageBoxImage.Information);
                                    }
                                }
                            }
                        }
                    }
                    }
                    else
                    {
                        Info("BŁĄD: Utworzenie dokumentu ZK zwróciło null.", "SubiektService");
                    }
                }
                catch (COMException comEx)
                {
                    Error($"Błąd COM podczas tworzenia dokumentu ZK: {comEx.Message}", "SubiektService");
                    // Wyczyść cache - instancja może być nieprawidłowa
                    _cachedSubiekt = null;
                    _cachedGt = null;
                    PowiadomOZmianieInstancji(false);
                    MessageBox.Show(
                        $"Błąd podczas tworzenia dokumentu ZK:\n\n{comEx.Message}\n\nSpróbuj ponownie - zostanie utworzona nowa instancja Subiekta GT.",
                        "Błąd",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    Error(ex, "SubiektService", "Błąd podczas tworzenia dokumentu ZK");
                    MessageBox.Show(
                        $"Dokument ZK został utworzony, ale wystąpił błąd podczas otwierania okna:\n\n{ex.Message}",
                        "Ostrzeżenie",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Critical(ex, "SubiektService", "Krytyczny błąd podczas otwierania okna ZK");
                System.Windows.MessageBox.Show(
                    $"Błąd podczas otwierania okna ZK:\n\n{ex.Message}",
                    "Błąd",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }
    }
}


