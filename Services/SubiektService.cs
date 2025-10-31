using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;

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
        /// Zamyka aktywną instancję Subiekta GT i zwalnia licencję
        /// </summary>
        public void ZwolnijLicencje()
        {
            try
            {
                if (_cachedSubiekt != null)
                {
                    Console.WriteLine("[SubiektService] Zamykanie instancji Subiekta GT i zwolnienie licencji...");
                    
                    try
                    {
                        // Próbuj zamknąć instancję przez Zamknij() jeśli dostępne
                        _cachedSubiekt.Zamknij();
                        Console.WriteLine("[SubiektService] Wywołano Zamknij() na instancji Subiekta GT.");
                    }
                    catch (Exception zamknijEx)
                    {
                        Console.WriteLine($"[SubiektService] Zamknij() nie zadziałał: {zamknijEx.Message}");
                        // Spróbuj innej metody zamknięcia
                        try
                        {
                            // Ustaw okno jako niewidoczne i spróbuj zamknąć
                            _cachedSubiekt.Okno.Widoczne = false;
                        }
                        catch
                        {
                            // Ignoruj błędy
                        }
                    }
                }
                
                // Wyczyść cache niezależnie od tego czy zamknięcie się powiodło
                _cachedSubiekt = null;
                _cachedGt = null;
                
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

        public void OtworzOknoZK(string? nip = null, System.Collections.Generic.IEnumerable<Gryzak.Models.Product>? items = null, double? couponAmount = null, double? subTotal = null, string? couponTitle = null, string? orderId = null, double? handlingAmount = null, double? shippingAmount = null)
        {
            try
            {
                Console.WriteLine($"[SubiektService] Próba otwarcia okna ZK{(nip != null ? $" z kontrahentem o NIP: {nip}" : " bez kontrahenta")}...");

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
                                }
                            }
                            catch (Exception kontrEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się wyszukać kontrahenta: {kontrEx.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("[SubiektService] Nie podano NIP - dokument ZK zostanie otwarty bez kontrahenta.");
                        }
                        
                        // Ustaw numer zamówienia jako numer oryginalnego dokumentu
                        if (!string.IsNullOrWhiteSpace(orderId))
                        {
                            try
                            {
                                zkDokument.NumerOryginalny = orderId;
                                Console.WriteLine($"[SubiektService] Ustawiono numer oryginalnego dokumentu: {orderId}");
                            }
                            catch (Exception numOrigEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się ustawić NumerOryginalny: {numOrigEx.Message}");
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
                                    try
                                    {
                                        zkDokument.Uwagi = fullNote;
                                        Console.WriteLine($"[SubiektService] Dodano informacje do uwag ZK (Uwagi): {fullNote}");
                                    }
                                    catch (Exception directEx)
                                    {
                                        Console.WriteLine($"[SubiektService] Uwagi nie zadziałały: {directEx.Message}");
                                        
                                        // Ostatni fallback: Opis
                                        try
                                        {
                                            zkDokument.Opis = fullNote;
                                            Console.WriteLine($"[SubiektService] Dodano informacje do uwag ZK (Opis): {fullNote}");
                                        }
                                        catch
                                        {
                                            Console.WriteLine("[SubiektService] Nie udało się ustawić żadnego pola uwag dla dokumentu ZK.");
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
                                dynamic pozycje = zkDokument.Pozycje;
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

                                        // Cena z API (netto)
                                        double? apiPriceNet = null;
                                        if (!string.IsNullOrWhiteSpace(it.Price))
                                        {
                                            try
                                            {
                                                var parsed = double.Parse(it.Price, System.Globalization.CultureInfo.InvariantCulture);
                                                // Odejmij procent kuponu od ceny (jeśli kupon istnieje)
                                                if (couponPercentage.HasValue)
                                                {
                                                    parsed = parsed * (1.0 - couponPercentage.Value / 100.0);
                                                }
                                                
                                                apiPriceNet = parsed;
                                            }
                                            catch { apiPriceNet = null; }
                                        }

                                        try
                                        {
                                            // Cena bazowa z kartoteki (ustawiana automatycznie przez Subiekta po dodaniu pozycji)
                                            double basePrice = 0.0;
                                            try { basePrice = Convert.ToDouble(pozycja.CenaNettoPrzedRabatem, System.Globalization.CultureInfo.InvariantCulture); } catch { basePrice = 0.0; }

                                            if (apiPriceNet.HasValue && basePrice > 0.0)
                                            {
                                                // Jeśli cena API jest niższa od bazowej, policz rabat procentowy
                                                if (apiPriceNet.Value < basePrice)
                                                {
                                                    var discount = (1.0 - (apiPriceNet.Value / basePrice)) * 100.0;
                                                    if (discount < 0) discount = 0;
                                                    if (discount > 99.99) discount = 99.99;
                                                    try { pozycja.RabatProcent = discount; } catch { }
                                                }
                                                else
                                                {
                                                    // Cena API >= bazowej: rabat 0, opcjonalnie podnieś bazę do ceny API
                                                    try { pozycja.RabatProcent = 0.0; } catch { }
                                                    // Jeżeli chcemy wymusić dokładnie cenę API przy cenie wyższej niż bazowa
                                                    try { pozycja.CenaNettoPrzedRabatem = apiPriceNet.Value; } catch { }
                                                }
                                            }
                                            else if (apiPriceNet.HasValue)
                                            {
                                                // Brak wiarygodnej ceny bazowej: ustaw po prostu cenę przed rabatem
                                                try { pozycja.CenaNettoPrzedRabatem = apiPriceNet.Value; } catch { }
                                            }
                                        }
                                        catch { }

                                        Console.WriteLine($"[SubiektService] Dodano pozycję towarową o ID={towarId} (qty={it.Quantity}, apiPrice={it.Price}).");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[SubiektService] Pominięto pozycję z nieprawidłowym product_id: '{it.ProductId}'");
                                    }
                                }
                            }
                            catch (Exception dodajPozEx)
                            {
                                Console.WriteLine($"[SubiektService] Nie udało się dodać pozycji: {dodajPozEx.Message}");
                            }
                        }
                        
                        

                        // Dodaj towar 'KOSZTY/1' na końcu
                        // Tylko jeśli w API jest handling
                        if (handlingAmount.HasValue)
                        {
                            try
                            {
                                dynamic towary = subiekt.Towary;
                                dynamic towarKoszty = towary.Wczytaj("KOSZTY/1");
                                
                                if (towarKoszty != null)
                                {
                                    dynamic pozycje = zkDokument.Pozycje;
                                    dynamic kosztyPoz = pozycje.Dodaj(towarKoszty.Identyfikator);
                                    try { kosztyPoz.IloscJm = 1; } catch { }
                                    // Ustaw cenę na wartość handling z API
                                    try 
                                    { 
                                        kosztyPoz.CenaNettoPrzedRabatem = handlingAmount.Value;
                                        Console.WriteLine($"[SubiektService] Dodano towar 'KOSZTY/1' z ceną handling ({handlingAmount.Value:F2}).");
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
                            Console.WriteLine("[SubiektService] Brak handling w API - pomijam dodawanie 'KOSZTY/1'.");
                        }
                        
                        // Dodaj towar 'KOSZTY/2' na końcu
                        // Tylko jeśli w API jest shipping
                        if (shippingAmount.HasValue)
                        {
                            try
                            {
                                dynamic towary = subiekt.Towary;
                                dynamic towarShipping = towary.Wczytaj("KOSZTY/2");
                                
                                if (towarShipping != null)
                                {
                                    dynamic pozycje = zkDokument.Pozycje;
                                    dynamic shippingPoz = pozycje.Dodaj(towarShipping.Identyfikator);
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

                        // Otwórz okno dokumentu używając metody Wyswietl() - zgodnie z przykładem
                        try
                        {
                            // Wyswietl(false) - otwiera okno w trybie edycji
                            zkDokument.Wyswietl(false);
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


