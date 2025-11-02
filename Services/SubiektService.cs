using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Gryzak.Views;
using Microsoft.Data.SqlClient;

namespace Gryzak.Services
{
    public class SubiektService
    {
        // Stała dla typu dokumentu ZK (zamówienie od klienta)
        private const int gtaSubiektDokumentZK = unchecked((int)0xFFFFFFF8); // -8
        
        // Cache dla obiektu GT i Subiekta - aby nie uruchamiać od nowa za każdym razem
        private static dynamic? _cachedGt = null;
        private static dynamic? _cachedSubiekt = null;

        private readonly ConfigService _configService;

        public SubiektService()
        {
            _configService = new ConfigService();
        }

        public bool CzyInstancjaAktywna()
        {
            return _cachedSubiekt != null && _cachedGt != null;
        }
        
        /// <summary>
        /// Wyszukuje kontrahenta przez zapytanie SQL do bazy MSSQL
        /// </summary>
        private int? WyszukajKontrahentaPrzezSQL(string? nip, string? email = null, string? customerName = null, string? phone = null, string? company = null, string? address = null)
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
                    Console.WriteLine("[SubiektService] Brak adresu serwera MSSQL w konfiguracji - pomijam wyszukiwanie przez SQL.");
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
                        Console.WriteLine("[SubiektService] Brak adresu email i nazwy klienta - pomijam wyszukiwanie przez SQL.");
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
                    
                    Console.WriteLine($"[SubiektService] Wykonuję zapytanie SQL do wyszukania kontrahenta (email: {email}, customerName: {customerName}):");
                    
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
                        Console.WriteLine($"[SubiektService] {loggedQuery}");
                        
                        
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
                                Console.WriteLine($"[SubiektService] Znaleziono {kontrahenci.Count} kontrahentów przez SQL.");
                            }
                            else
                            {
                                Console.WriteLine($"[SubiektService] Nie znaleziono kontrahentów przez SQL - wyświetlam dialog z pustą listą.");
                            }
                            
                            Console.WriteLine($"[SubiektService] Wyświetlam dialog wyboru kontrahenta ({kontrahenci.Count} wyników)...");
                            
                            // Użyj synchronicznego Invoke, aby upewnić się że dialog jest wyświetlony
                            Application.Current?.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    Console.WriteLine("[SubiektService] Tworzę dialog SelectKontrahentDialog...");
                                    var dialog = new SelectKontrahentDialog(kontrahenci, customerName, email, phone, company, nip, address);
                                    
                                    // Ustaw właściciela dialogu, aby był widoczny
                                    if (Application.Current?.MainWindow != null)
                                    {
                                        dialog.Owner = Application.Current.MainWindow;
                                        dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                                    }
                                    else
                                    {
                                        dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                                    }
                                    
                                    // Ustaw dialog na wierzchu, aby był widoczny
                                    dialog.Topmost = true;
                                    
                                    Console.WriteLine("[SubiektService] Wyświetlam dialog ShowDialog()...");
                                    bool? result = dialog.ShowDialog();
                                    
                                    // Wyłącz Topmost po zamknięciu dialogu
                                    dialog.Topmost = false;
                                    Console.WriteLine($"[SubiektService] Dialog ShowDialog() zakończył się z wynikiem: {result}");
                                    
                                    if (result == true && dialog.SelectedKontrahent != null)
                                    {
                                        selectedId = dialog.SelectedKontrahent.Id;
                                        Console.WriteLine($"[SubiektService] Użytkownik wybrał kontrahenta: ID={dialog.SelectedKontrahent.Id}, Symbol={dialog.SelectedKontrahent.Symbol}");
                                    }
                                    else
                                    {
                                        Console.WriteLine("[SubiektService] Użytkownik anulował wybór kontrahenta lub nie wybrał żadnego - ZK zostanie otwarte bez kontrahenta.");
                                    }
                                }
                                catch (Exception dialogEx)
                                {
                                    Console.WriteLine($"[SubiektService] Błąd podczas wyświetlania dialogu wyboru kontrahenta: {dialogEx.Message}");
                                    Console.WriteLine($"[SubiektService] Stack trace: {dialogEx.StackTrace}");
                                }
                            }, System.Windows.Threading.DispatcherPriority.Normal);
                            
                            Console.WriteLine($"[SubiektService] WyszukajKontrahentaPrzezSQL zwraca: {selectedId?.ToString() ?? "null"}");
                            return selectedId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SubiektService] Błąd podczas wyszukiwania kontrahenta przez SQL: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Zamyka aktywną instancję Subiekta GT i zwalnia licencję
        /// </summary>
        public void ZwolnijLicencje()
        {
            try
            {
                Console.WriteLine("[SubiektService] Zamykanie instancji Subiekta GT i zwolnienie licencji...");
                
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
                                Console.WriteLine($"[SubiektService] Wywołano Zakoncz() na obiekcie Aplikacja. Wynik: {zakonczWynik}");
                                
                                // Poczekaj chwilę na zamknięcie procesu
                                System.Threading.Thread.Sleep(500);
                            }
                        }
                        catch (Exception aplikacjaEx)
                        {
                            Console.WriteLine($"[SubiektService] Nie udało się zamknąć przez Aplikacja.Zakoncz(): {aplikacjaEx.Message}");
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
                                            Console.WriteLine($"[SubiektService] Wywołano Zakoncz() na obiekcie GT.Aplikacja. Wynik: {zakonczWynik}");
                                            System.Threading.Thread.Sleep(500);
                                        }
                                    }
                                    catch
                                    {
                                        Console.WriteLine("[SubiektService] Nie można uzyskać dostępu do GT.Aplikacja");
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
                        Console.WriteLine($"[SubiektService] Błąd podczas próby zamknięcia aplikacji: {zamknijEx.Message}");
                    }
                    
                    // Zwolnij referencje COM obiektów
                    try
                    {
                        if (_cachedSubiekt != null)
                        {
                            Marshal.ReleaseComObject(_cachedSubiekt);
                            Console.WriteLine("[SubiektService] Zwolniono referencję COM dla obiektu Subiekt.");
                        }
                    }
                    catch (Exception releaseEx)
                    {
                        Console.WriteLine($"[SubiektService] Błąd podczas zwalniania obiektu Subiekt: {releaseEx.Message}");
                    }
                }
                
                if (_cachedGt != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(_cachedGt);
                        Console.WriteLine("[SubiektService] Zwolniono referencję COM dla obiektu GT.");
                    }
                    catch (Exception releaseEx)
                    {
                        Console.WriteLine($"[SubiektService] Błąd podczas zwalniania obiektu GT: {releaseEx.Message}");
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
                
                Console.WriteLine("[SubiektService] Instancja Subiekta GT została zamknięta, licencja zwolniona.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SubiektService] BŁĄD podczas zamykania instancji: {ex.Message}");
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

        public void OtworzOknoZK(string? nip = null, System.Collections.Generic.IEnumerable<Gryzak.Models.Product>? items = null, double? couponAmount = null, double? subTotal = null, string? couponTitle = null, string? orderId = null, double? handlingAmount = null, double? shippingAmount = null, string? currency = null, double? codFeeAmount = null, string? orderTotal = null, double? glsAmount = null, string? email = null, string? customerName = null, string? phone = null, string? company = null, string? address = null)
        {
            try
            {
                Console.WriteLine($"[SubiektService] Próba otwarcia okna ZK{(nip != null ? $" z kontrahentem o NIP: {nip}" : " bez kontrahenta")}{(email != null ? $" (email: {email})" : "")}...");

                dynamic? gt = null;
                dynamic? subiekt = null;

                // Sprawdź czy mamy już uruchomioną instancję w cache
                // Używamy prostej weryfikacji - jeśli obiekt jest w cache, zakładamy że działa
                // Jeśli nie działa, catch podczas użycia złapie wyjątek i wtedy wyczyścimy cache
                if (_cachedSubiekt != null && _cachedGt != null)
                {
                    subiekt = _cachedSubiekt;
                    gt = _cachedGt;
                    Console.WriteLine("[SubiektService] Używam istniejącej instancji Subiekta GT z cache.");
                    // Upewnij się, że status jest aktywny (może nie zostać ustawiony przy starcie aplikacji)
                    PowiadomOZmianieInstancji(true);
                }

                // Jeśli nie ma działającej instancji, utwórz nową
                if (subiekt == null)
                {
                    Console.WriteLine("[SubiektService] Uruchamiam nową instancję Subiekta GT...");

                    // Utworz obiekt GT (COM)
                    Type? gtType = Type.GetTypeFromProgID("InsERT.gt");
                    if (gtType == null)
                    {
                        Console.WriteLine("[SubiektService] BŁĄD: Nie można załadować typu COM 'InsERT.gt'. Upewnij się, że Sfera jest zainstalowana.");
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
                        Console.WriteLine("[SubiektService] BŁĄD: Nie można utworzyć instancji obiektu GT.");
                        return;
                    }

                    // Ustaw parametry połączenia i logowania
                    try
                    {
                        gt.Produkt = 1; // gtaProduktSubiekt
                        
                        // Wczytaj konfigurację i ustaw operatora i hasło PRZED uruchomieniem, aby pominąć okno logowania
                        var subiektConfig = _configService.LoadSubiektConfig();
                        gt.Operator = subiektConfig.User;
                        gt.OperatorHaslo = subiektConfig.Password;
                        
                        Console.WriteLine($"[SubiektService] Ustawiono operatora: {subiektConfig.User} - okno logowania zostanie pominięte.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SubiektService] BŁĄD podczas konfiguracji GT: {ex.Message}");
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
                            Console.WriteLine("[SubiektService] BŁĄD: Uruchomienie Subiekta GT zwróciło null.");
                            return;
                        }
                        Console.WriteLine("[SubiektService] Subiekt GT uruchomiony w tle (bez interfejsu użytkownika).");
                        

                        
                        // Zapisz w cache dla następnych użyć
                        _cachedGt = gt;
                        _cachedSubiekt = subiekt;
                        
                        // Powiadom o aktywności instancji
                        PowiadomOZmianieInstancji(true);
                        
                        // Upewnij się, że główne okno jest ukryte (dodatkowa ochrona)
                        try
                        {
                            subiekt.Okno.Widoczne = false;
                            Console.WriteLine("[SubiektService] Główne okno Subiekta GT ustawione jako niewidoczne.");
                        }
                        catch (Exception oknoEx)
                        {
                            Console.WriteLine($"[SubiektService] Uwaga: Nie można ustawić głównego okna jako niewidoczne: {oknoEx.Message}");
                        }
                    }
                        catch (COMException comEx)
                    {
                        Console.WriteLine($"[SubiektService] BŁĄD COM: {comEx.Message} (HRESULT: 0x{comEx.ErrorCode:X8})");
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
                        Console.WriteLine($"[SubiektService] BŁĄD: {ex.Message}");
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

                // Dodaj dokument ZK i otwórz jego okno
                // Używamy zmiennej subiekt, która może pochodzić z cache lub być świeżo utworzona
                try
                {
                    Console.WriteLine("[SubiektService] Próba dodania dokumentu ZK...");
                    dynamic? zkDokument = null;

                    // Spróbuj użyć SuDokumentyManager.DodajZK() - to jest zalecana metoda
                    try
                    {
                        dynamic dokumentyManager = subiekt.SuDokumentyManager;
                        if (dokumentyManager != null)
                        {
                            zkDokument = dokumentyManager.DodajZK();
                            Console.WriteLine("[SubiektService] Użyto SuDokumentyManager.DodajZK()");
                        }
                    }
                    catch (Exception ex1)
                    {
                        Console.WriteLine($"[SubiektService] SuDokumentyManager.DodajZK() nie zadziałał: {ex1.Message}");
                    }

                    // Fallback: użyj Dokumenty.DodajZK()
                    if (zkDokument == null)
                    {
                        try
                        {
                            dynamic dokumenty = subiekt.Dokumenty;
                            zkDokument = dokumenty.DodajZK();
                            Console.WriteLine("[SubiektService] Użyto Dokumenty.DodajZK()");
                        }
                        catch (Exception ex2)
                        {
                            Console.WriteLine($"[SubiektService] Dokumenty.DodajZK() nie zadziałał: {ex2.Message}");
                        }
                    }

                    // Ostatni fallback: użyj Dokumenty.Dodaj() z parametrem
                    if (zkDokument == null)
                    {
                        try
                        {
                            dynamic dokumenty = subiekt.Dokumenty;
                            zkDokument = dokumenty.Dodaj(gtaSubiektDokumentZK);
                            Console.WriteLine("[SubiektService] Użyto Dokumenty.Dodaj(gtaSubiektDokumentZK)");
                        }
                        catch (Exception ex3)
                        {
                            Console.WriteLine($"[SubiektService] Dokumenty.Dodaj() nie zadziałał: {ex3.Message}");
                        }
                    }

                    if (zkDokument != null)
                    {
                        Console.WriteLine("[SubiektService] Dokument ZK został utworzony.");
                        
                        // Wyszukaj kontrahenta po NIP i ustaw w dokumencie (tylko jeśli NIP został podany)
                        if (!string.IsNullOrWhiteSpace(nip))
                        {
                            try
                            {
                                // Pobierz menedżer kontrahentów
                                dynamic kontrahenciManager = subiekt.KontrahenciManager;
                                
                                // Wyszukaj kontrahenta po NIP - gtaKontrahentWgNip = 2
                                dynamic kontrahent = kontrahenciManager.WczytajKontrahentaWg(nip, 2);
                                
                                if (kontrahent != null)
                                {
                                    // Ustaw ID kontrahenta w dokumencie
                                    zkDokument.KontrahentId = kontrahent.Identyfikator;
                                    Console.WriteLine($"[SubiektService] Ustawiono kontrahenta o ID={kontrahent.Identyfikator} (NIP: {nip}) w dokumencie ZK.");
                                }
                                else
                                {
                                    Console.WriteLine($"[SubiektService] Nie znaleziono kontrahenta o NIP: {nip}");
                                    
                                    // Spróbuj wyszukać kontrahenta przez SQL po emailu
                                    Console.WriteLine($"[SubiektService] Próba wyszukania kontrahenta przez SQL...");
                                    int? kontrahentId = WyszukajKontrahentaPrzezSQL(nip, email, customerName, phone, company, address);
                                    
                                    if (kontrahentId.HasValue)
                                    {
                                        // Ustaw ID kontrahenta w dokumencie
                                        zkDokument.KontrahentId = kontrahentId.Value;
                                        Console.WriteLine($"[SubiektService] Ustawiono kontrahenta o ID={kontrahentId.Value} (znaleziony przez SQL) w dokumencie ZK.");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[SubiektService] Nie znaleziono kontrahenta przez SQL.");
                                    }
                                }
                            }
                            catch (Exception kontrEx)
                            {
                                Console.WriteLine($"[SubiektService] BŁĄD: Nie udało się wyszukać kontrahenta: {kontrEx.Message}");
                                Console.WriteLine($"[SubiektService] Stack trace: {kontrEx.StackTrace}");
                                
                                // Próbuj wyszukać przez SQL nawet jeśli wystąpił błąd
                                try
                                {
                                    Console.WriteLine($"[SubiektService] Próba wyszukania kontrahenta przez SQL (po błędzie NIP)...");
                                    int? kontrahentId = WyszukajKontrahentaPrzezSQL(nip, email, customerName, phone, company, address);
                                    
                                    if (kontrahentId.HasValue && zkDokument != null)
                                    {
                                        // Ustaw ID kontrahenta w dokumencie
                                        int id = kontrahentId.GetValueOrDefault();
                                        zkDokument!.KontrahentId = id;
                                        Console.WriteLine($"[SubiektService] Ustawiono kontrahenta o ID={id} (znaleziony przez SQL po błędzie) w dokumencie ZK.");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[SubiektService] Nie znaleziono kontrahenta przez SQL po błędzie.");
                                    }
                                }
                                catch (Exception sqlEx)
                                {
                                    Console.WriteLine($"[SubiektService] Błąd podczas wyszukiwania przez SQL po błędzie NIP: {sqlEx.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("[SubiektService] Nie podano NIP - próba wyszukania kontrahenta przez SQL...");
                            
                            // Próbuj wyszukać kontrahenta przez SQL nawet gdy nie ma NIP
                            try
                            {
                                int? kontrahentId = WyszukajKontrahentaPrzezSQL(null, email, customerName, phone, company, address);
                                
                                if (kontrahentId.HasValue && zkDokument != null)
                                {
                                    // Ustaw ID kontrahenta w dokumencie
                                    int id = kontrahentId.GetValueOrDefault();
                                    zkDokument!.KontrahentId = id;
                                    Console.WriteLine($"[SubiektService] Ustawiono kontrahenta o ID={id} (znaleziony przez SQL bez NIP) w dokumencie ZK.");
                                }
                                else
                                {
                                    Console.WriteLine("[SubiektService] Nie znaleziono kontrahenta przez SQL - dokument ZK zostanie otwarty bez kontrahenta.");
                                }
                            }
                            catch (Exception sqlEx)
                            {
                                Console.WriteLine($"[SubiektService] Błąd podczas wyszukiwania przez SQL (bez NIP): {sqlEx.Message}");
                                Console.WriteLine("[SubiektService] Dokument ZK zostanie otwarty bez kontrahenta.");
                            }
                        }
                        
                        // Ustaw numer zamówienia jako numer oryginalnego dokumentu
                        if (!string.IsNullOrWhiteSpace(orderId) && zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.NumerOryginalny = orderId;
                                Console.WriteLine($"[SubiektService] Ustawiono numer oryginalnego dokumentu: {orderId}");
                            }
                            catch (Exception numOrigEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się ustawić NumerOryginalny: {numOrigEx.Message}");
                            }
                        }
                        
                        // Ustaw walutę dokumentu (jeśli została podana)
                        // Używamy atrybutu WalutaSymbol, który przyjmuje symbol waluty jako string (np. "EUR", "USD", "PLN")
                        if (!string.IsNullOrWhiteSpace(currency) && zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.WalutaSymbol = currency;
                                Console.WriteLine($"[SubiektService] Ustawiono walutę dokumentu ZK: {currency}");
                                
                                // Opcjonalnie można pobrać kurs waluty automatycznie po ustawieniu symbolu
                                try
                                {
                                    zkDokument!.PobierzKursWaluty();
                                    Console.WriteLine($"[SubiektService] Pobrano kurs waluty dla {currency}");
                                }
                                catch (Exception kursEx)
                                {
                                    Console.WriteLine($"[SubiektService] Uwaga: Nie udało się pobrać kursu waluty automatycznie: {kursEx.Message}");
                                    // To nie jest błąd krytyczny, kurs można ustawić ręcznie później
                                }
                            }
                            catch (Exception walutaEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się ustawić waluty: {walutaEx.Message}");
                            }
                        }
                        
                        // Ustaw przeliczanie dokumentu od cen brutto (ponieważ ustawiamy CenaBruttoPoRabacie dla pozycji)
                        if (zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.LiczonyOdCenBrutto = true;
                                Console.WriteLine("[SubiektService] Ustawiono przeliczanie dokumentu ZK od cen brutto (LiczoneOdCenBrutto = true)");
                            }
                            catch (Exception bruttoEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się ustawić LiczoneOdCenBrutto: {bruttoEx.Message}");
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
                                    notes.Add($"{couponTitle} ({couponAmount.Value:F2} zł)");
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
                                        Console.WriteLine($"[SubiektService] Dodano informacje do uwag ZK (ZmienOpisDokumentu): {fullNote}");
                                    }
                                }
                                catch (Exception methodEx)
                                {
                                    Console.WriteLine($"[SubiektService] ZmienOpisDokumentu nie zadziałała: {methodEx.Message}");
                                    
                                    // Fallback: spróbuj bezpośrednio ustawić Uwagi
                                    if (zkDokument != null)
                                    {
                                        try
                                        {
                                            zkDokument!.Uwagi = fullNote;
                                            Console.WriteLine($"[SubiektService] Dodano informacje do uwag ZK (Uwagi): {fullNote}");
                                        }
                                        catch (Exception directEx)
                                        {
                                            Console.WriteLine($"[SubiektService] Uwagi nie zadziałały: {directEx.Message}");
                                            
                                            // Ostatni fallback: Opis
                                            try
                                            {
                                                zkDokument!.Opis = fullNote;
                                                Console.WriteLine($"[SubiektService] Dodano informacje do uwag ZK (Opis): {fullNote}");
                                            }
                                            catch
                                            {
                                                Console.WriteLine("[SubiektService] Nie udało się ustawić żadnego pola uwag dla dokumentu ZK.");
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception couponEx)
                            {
                                Console.WriteLine($"[SubiektService] Błąd dodawania informacji do uwag ZK: {couponEx.Message}");
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
                                Console.WriteLine($"[SubiektService] ANALIZA KUPONU:");
                                Console.WriteLine($"[SubiektService] Wartość kuponu: {couponAmount.Value:F2}");
                                Console.WriteLine($"[SubiektService] Suma wartości produktów: {totalProductsValue:F2}");
                                Console.WriteLine($"[SubiektService] Procent kuponu względem produktów: {couponPercentage:F2}%");
                            }
                            else
                            {
                                Console.WriteLine("[SubiektService] Brak danych o cenach produktów - nie można obliczyć procentu kuponu.");
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
                                                Console.WriteLine($"[SubiektService] Błąd parsowania ceny dla produktu {it.ProductId}: {priceEx.Message}");
                                                apiPriceBrutto = null;
                                            }
                                        }

                                        // Ustaw cenę brutto po rabacie z API
                                        if (apiPriceBrutto.HasValue)
                                        {
                                            try
                                            {
                                                pozycja.CenaBruttoPoRabacie = apiPriceBrutto.Value;
                                                Console.WriteLine($"[SubiektService] Ustawiono CenaBruttoPoRabacie={apiPriceBrutto.Value:F2} dla produktu ID={towarId} (Price={it.Price}, Tax={it.Tax})");
                                            }
                                            catch (Exception cenaEx)
                                            {
                                                Console.WriteLine($"[SubiektService] Nie udało się ustawić CenaBruttoPoRabacie: {cenaEx.Message}");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[SubiektService] Brak ceny brutto z API dla produktu ID={towarId}");
                                        }
                                        
                                        Console.WriteLine($"[SubiektService] Dodano pozycję towarową o ID={towarId} (qty={it.Quantity}, apiPriceNetto={it.Price}, apiPriceBrutto={apiPriceBrutto?.ToString("F2") ?? "brak"}).");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[SubiektService] Pominięto pozycję z nieprawidłowym product_id: '{it.ProductId}'");
                                    }
                                    }
                                }
                            }
                            catch (Exception dodajPozEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się dodać pozycji: {dodajPozEx.Message}");
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
                                        Console.WriteLine($"[SubiektService] Dodano towar 'KOSZTY/1' z ceną {sumaKosztowKOSZTY1:F2} (składniki: {string.Join(" + ", skladniki)}).");
                                    }
                                    catch
                                    {
                                        Console.WriteLine("[SubiektService] Dodano towar 'KOSZTY/1' na końcu pozycji, ale nie udało się ustawić ceny.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("[SubiektService] Nie znaleziono towaru o symbolu 'KOSZTY/1'.");
                                }
                            }
                            catch (Exception kosztyEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się dodać towaru 'KOSZTY/1': {kosztyEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("[SubiektService] Brak handling i gls w API - pomijam dodawanie 'KOSZTY/1'.");
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
                                        Console.WriteLine($"[SubiektService] Dodano towar 'KOSZTY/2' z ceną shipping ({shippingAmount.Value:F2}).");
                                    }
                                    catch
                                    {
                                        Console.WriteLine("[SubiektService] Dodano towar 'KOSZTY/2' na końcu pozycji, ale nie udało się ustawić ceny.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("[SubiektService] Nie znaleziono towaru o symbolu 'KOSZTY/2'.");
                                }
                            }
                            catch (Exception shippingEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się dodać towaru 'KOSZTY/2': {shippingEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("[SubiektService] Brak shipping w API - pomijam dodawanie 'KOSZTY/2'.");
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
                                        Console.WriteLine($"[SubiektService] Dodano towar 'KOSZTY/2' z ceną cod_fee ({codFeeAmount.Value:F2}).");
                                    }
                                    catch
                                    {
                                        Console.WriteLine("[SubiektService] Dodano towar 'KOSZTY/2' na końcu pozycji, ale nie udało się ustawić ceny.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("[SubiektService] Nie znaleziono towaru o symbolu 'KOSZTY/2'.");
                                }
                            }
                            catch (Exception codFeeEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się dodać towaru 'KOSZTY/2' dla cod_fee: {codFeeEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("[SubiektService] Brak cod_fee w API - pomijam dodawanie 'KOSZTY/2' dla cod_fee.");
                        }
                        
                        // Przelicz dokument przed odczytem wartości (może być wymagane)
                        if (zkDokument != null)
                        {
                            try
                            {
                                zkDokument!.Przelicz();
                                Console.WriteLine("[SubiektService] Dokument ZK został przeliczony przed odczytem wartości pozycji.");
                            }
                            catch (Exception przeliczEx)
                            {
                                Console.WriteLine($"[SubiektService] Uwaga: Nie udało się przeliczyć dokumentu przed odczytem wartości: {przeliczEx.Message}");
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
                            Console.WriteLine($"[SubiektService] Liczba pozycji w dokumencie ZK: {liczbaPozycji}");
                            
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
                                        Console.WriteLine($"[SubiektService] Błąd odczytu WartoscBruttoPoRabacie dla pozycji {i}: {wartoscEx.Message}");
                                    }
                                    
                                    sumaWartosci += wartoscPozycji;
                                    Console.WriteLine($"[SubiektService] Pozycja {i}: Wartość = {wartoscPozycji:F2}");
                                    
                                    // Dodatkowe logowanie - sprawdź jakie właściwości pozycji są dostępne
                                    try
                                    {
                                        var ilosc = pozycja.IloscJm;
                                        Console.WriteLine($"[SubiektService]   - Ilość: {ilosc}");
                                    }
                                    catch { }
                                    try
                                    {
                                        var cenaBruttoPoRabacie = pozycja.CenaBruttoPoRabacie;
                                        Console.WriteLine($"[SubiektService]   - CenaBruttoPoRabacie: {cenaBruttoPoRabacie:F2}");
                                    }
                                    catch { }
                                    try
                                    {
                                        var cenaNettoPoRabacie = pozycja.CenaNettoPoRabacie;
                                        Console.WriteLine($"[SubiektService]   - CenaNettoPoRabacie: {cenaNettoPoRabacie:F2}");
                                    }
                                    catch { }
                                }
                                catch (Exception pozEx)
                                {
                                    Console.WriteLine($"[SubiektService] Błąd odczytu pozycji {i}: {pozEx.Message}");
                                }
                            }
                            
                            Console.WriteLine($"[SubiektService] =========================================");
                            Console.WriteLine($"[SubiektService] SUMA WARTOŚCI WSZYSTKICH POZYCJI: {sumaWartosci:F2}");
                            Console.WriteLine($"[SubiektService] =========================================");
                            
                            // Porównaj sumę pozycji ZK z wartością zamówienia z API
                            if (!string.IsNullOrWhiteSpace(orderTotal))
                            {
                                try
                                {
                                    if (double.TryParse(orderTotal, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double orderTotalValue))
                                    {
                                        double roznica = Math.Abs(sumaWartosci - orderTotalValue);
                                        double roznicaProcent = sumaWartosci > 0 ? (roznica / sumaWartosci) * 100.0 : 0.0;
                                        
                                        Console.WriteLine($"[SubiektService] Wartość zamówienia z API: {orderTotalValue:F2}");
                                        Console.WriteLine($"[SubiektService] Różnica: {roznica:F2} ({roznicaProcent:F2}%)");
                                        
                                        // Jeśli różnica jest większa niż próg tolerancji (0.001 = 0.01 grosza)
                                        // Używamy progu zamiast > 0.0 aby uniknąć błędów zaokrąglenia zmiennoprzecinkowych
                                        if (roznica > 0.001)
                                        {
                                            // Analiza przyczyn różnicy
                                            Console.WriteLine($"[SubiektService] =========================================");
                                            Console.WriteLine($"[SubiektService] ANALIZA RÓŻNICY W WARTOŚCIACH");
                                            Console.WriteLine($"[SubiektService] =========================================");
                                            
                                            var wnioski = new System.Collections.Generic.List<string>();
                                            
                                            // Porównaj bezpośrednio sumę ZK z wartością z API
                                            // Wartość z API jest już finalną wartością (uwzględnia wszystkie składniki)
                                            Console.WriteLine($"[SubiektService] Suma pozycji ZK: {sumaWartosci:F2}");
                                            Console.WriteLine($"[SubiektService] Wartość zamówienia z API (total): {orderTotalValue:F2}");
                                            Console.WriteLine($"[SubiektService] Różnica: {roznica:F2} ({roznicaProcent:F2}%)");
                                            
                                            // Sprawdź komponenty zamówienia z API (dla informacji)
                                            double sumaKosztow = 0.0;
                                            if (handlingAmount.HasValue && handlingAmount.Value > 0.01)
                                            {
                                                sumaKosztow += handlingAmount.Value;
                                                Console.WriteLine($"[SubiektService] Handling z API: {handlingAmount.Value:F2}");
                                            }
                                            if (glsAmount.HasValue && glsAmount.Value > 0.01)
                                            {
                                                sumaKosztow += glsAmount.Value;
                                                Console.WriteLine($"[SubiektService] Gls z API: {glsAmount.Value:F2}");
                                            }
                                            if (shippingAmount.HasValue && shippingAmount.Value > 0.01)
                                            {
                                                sumaKosztow += shippingAmount.Value;
                                                Console.WriteLine($"[SubiektService] Shipping z API: {shippingAmount.Value:F2}");
                                            }
                                            if (codFeeAmount.HasValue && codFeeAmount.Value > 0.01)
                                            {
                                                sumaKosztow += codFeeAmount.Value;
                                                Console.WriteLine($"[SubiektService] Cod_fee z API: {codFeeAmount.Value:F2}");
                                            }
                                            if (sumaKosztow > 0.01)
                                            {
                                                Console.WriteLine($"[SubiektService] Suma kosztów z API: {sumaKosztow:F2}");
                                            }
                                            
                                            if (couponAmount.HasValue && couponAmount.Value > 0.01)
                                            {
                                                Console.WriteLine($"[SubiektService] Kupon z API: -{couponAmount.Value:F2}");
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
                                            Console.WriteLine($"[SubiektService] =========================================");
                                            Console.WriteLine($"[SubiektService] WNIOSKI:");
                                            foreach (var wniosek in wnioski)
                                            {
                                                Console.WriteLine($"[SubiektService]   - {wniosek}");
                                            }
                                            Console.WriteLine($"[SubiektService] =========================================");
                                            
                                            string komunikat = $"Suma wartości pozycji w dokumencie ZK: {sumaWartosci:F2} zł\nWartość zamówienia z API: {orderTotalValue:F2} zł\n\nRóżnica: {roznica:F2} zł ({roznicaProcent:F2}%)\n\nKorekta zostanie wykonana w następującej kolejności:\n1. Pozycja KOSZTY/1 (jeśli istnieje)\n2. Pozycja KOSZTY/2 (jeśli istnieje)\n3. Pierwsza pozycja produktowa";
                                            
                                            var dialog = new Gryzak.Views.KorektaWartosciDialog(komunikat);
                                            bool czyKorygowac = dialog.ShowDialog() == true && dialog.CzyKorygowac;
                                            
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
                                                            Console.WriteLine("[SubiektService] Nie udało się odczytać identyfikatorów towarów KOSZTY.");
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
                                                                            
                                                                            Console.WriteLine($"[SubiektService] Skorygowano pozycję KOSZTY/1 (pozycja {i}):");
                                                                            Console.WriteLine($"[SubiektService]   Stara wartość: {aktualnaWartosc:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Nowa wartość: {nowaWartosc:F2}");
                                                                            
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
                                                                            
                                                                            Console.WriteLine($"[SubiektService] Skorygowano pozycję KOSZTY/2 (pozycja {i}):");
                                                                            Console.WriteLine($"[SubiektService]   Stara wartość: {aktualnaWartosc:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Nowa wartość: {nowaWartosc:F2}");
                                                                            
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
                                                                            
                                                                            Console.WriteLine($"[SubiektService] Skorygowano pierwszą pozycję produktową (pozycja {i}):");
                                                                            Console.WriteLine($"[SubiektService]   Stara wartość: {aktualnaWartosc:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Różnica: {(czyZKwieksze ? "-" : "+")}{pozostalaRoznica:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Nowa wartość: {nowaWartosc:F2}");
                                                                            Console.WriteLine($"[SubiektService]   Nowa cena brutto: {nowaCenaBrutto:F2}");
                                                                            
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
                                                        Console.WriteLine($"[SubiektService] Suma po korekcie: {sumaPoKorekcie:F2}");
                                                        Console.WriteLine($"[SubiektService] Różnica po korekcie: {roznicaPoKorekcie:F2}");
                                                        
                                                        if (roznicaPoKorekcie <= 0.001)
                                                        {
                                                            Console.WriteLine($"[SubiektService] Korekta zakończona pomyślnie. Wartości są teraz zgodne.");
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
                                                        Console.WriteLine("[SubiektService] Błąd: Brak pozycji w dokumencie - nie można wykonać korekty.");
                                                        MessageBox.Show(
                                                            "Nie można skorygować - brak pozycji w dokumencie.",
                                                            "Błąd korekty",
                                                            System.Windows.MessageBoxButton.OK,
                                                            System.Windows.MessageBoxImage.Error);
                                                    }
                                                }
                                                catch (Exception cenaEx)
                                                {
                                                    Console.WriteLine($"[SubiektService] Błąd podczas korekty ceny: {cenaEx.Message}");
                                                    MessageBox.Show(
                                                        $"Błąd podczas korekty: {cenaEx.Message}",
                                                        "Błąd korekty",
                                                        System.Windows.MessageBoxButton.OK,
                                                        System.Windows.MessageBoxImage.Error);
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine("[SubiektService] Użytkownik wybrał 'Pozostaw' - różnica nie została skorygowana.");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[SubiektService] Wartości się zgadzają (różnica = 0.00).");
                                        }
                                    }
                                }
                                catch (Exception porownanieEx)
                                {
                                    Console.WriteLine($"[SubiektService] Błąd podczas porównywania wartości: {porownanieEx.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[SubiektService] Brak wartości zamówienia z API do porównania.");
                            }
                            
                            // Spróbuj też odczytać wartość bezpośrednio z dokumentu (jeśli dostępna)
                            if (zkDokument != null)
                            {
                                try
                                {
                                    double wartoscDokumentuBrutto = zkDokument.WartoscBrutto;
                                    Console.WriteLine($"[SubiektService] Wartość brutto dokumentu ZK (z właściwości): {wartoscDokumentuBrutto:F2}");
                                }
                                catch { }
                            }
                            
                            if (zkDokument != null)
                            {
                                try
                                {
                                    double wartoscDokumentuNetto = zkDokument!.WartoscNetto;
                                    Console.WriteLine($"[SubiektService] Wartość netto dokumentu ZK (z właściwości): {wartoscDokumentuNetto:F2}");
                                }
                                catch { }
                                
                                try
                                {
                                    double wartoscDokumentu = zkDokument!.Wartosc;
                                    Console.WriteLine($"[SubiektService] Wartość dokumentu ZK (z właściwości): {wartoscDokumentu:F2}");
                                }
                                catch { }
                            }
                        }
                        catch (Exception sumaEx)
                        {
                            Console.WriteLine($"[SubiektService] Błąd podczas odczytu sumy wartości pozycji: {sumaEx.Message}");
                        }

                        // Otwórz okno dokumentu używając metody Wyswietl() - zgodnie z przykładem
                        if (zkDokument != null)
                        {
                            try
                            {
                                // Wyswietl(false) - otwiera okno w trybie edycji
                                zkDokument!.Wyswietl(false);
                                Console.WriteLine("[SubiektService] Wywołano Wyswietl(false) na dokumencie ZK.");

                                // Aktywuj okno ZK na wierzch - używamy głównego okna Subiekta
                                try
                                {
                                    subiekt.Okno.Aktywuj();
                                    Console.WriteLine("[SubiektService] Aktywowano główne okno Subiekta GT na wierzch.");
                                }
                                catch (Exception aktywujEx)
                                {
                                    Console.WriteLine($"[SubiektService] Aktywacja głównego okna Subiekta nie zadziałała: {aktywujEx.Message}");
                                }
                                
                                Console.WriteLine("[SubiektService] Okno dokumentu ZK zostało otwarte pomyślnie.");
                            }
                            catch (Exception wyswietlEx)
                            {
                                Console.WriteLine($"[SubiektService] BŁĄD: Wyswietl() nie zadziałał: {wyswietlEx.Message}");
                                
                                // Spróbuj bez parametru
                                if (zkDokument != null)
                                {
                                    try
                                    {
                                        zkDokument.Wyswietl();
                                        Console.WriteLine("[SubiektService] Wyswietl() bez parametru zadziałał.");
                                        
                                        // Aktywuj główne okno Subiekta na wierzch
                                        try
                                        {
                                            subiekt.Okno.Aktywuj();
                                            Console.WriteLine("[SubiektService] Aktywowano główne okno Subiekta na wierzch (fallback).");
                                        }
                                        catch (Exception aktywujEx2)
                                        {
                                            Console.WriteLine($"[SubiektService] Aktywacja głównego okna Subiekta (fallback) nie zadziałała: {aktywujEx2.Message}");
                                        }
                                    }
                                    catch (Exception wyswietl2Ex)
                                    {
                                        Console.WriteLine($"[SubiektService] BŁĄD: Wyswietl() bez parametru też nie zadziałał: {wyswietl2Ex.Message}");
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
                        Console.WriteLine("[SubiektService] BŁĄD: Utworzenie dokumentu ZK zwróciło null.");
                    }
                }
                catch (COMException comEx)
                {
                    Console.WriteLine($"[SubiektService] BŁĄD COM podczas tworzenia dokumentu ZK: {comEx.Message}");
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
                    Console.WriteLine($"[SubiektService] BŁĄD podczas tworzenia dokumentu ZK: {ex.Message}");
                    Console.WriteLine($"[SubiektService] Stack trace: {ex.StackTrace}");
                    MessageBox.Show(
                        $"Dokument ZK został utworzony, ale wystąpił błąd podczas otwierania okna:\n\n{ex.Message}",
                        "Ostrzeżenie",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SubiektService] KRYTYCZNY BŁĄD: {ex.Message}");
                Console.WriteLine($"[SubiektService] Stack trace: {ex.StackTrace}");
                System.Windows.MessageBox.Show(
                    $"Błąd podczas otwierania okna ZK:\n\n{ex.Message}",
                    "Błąd",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }
    }
}


