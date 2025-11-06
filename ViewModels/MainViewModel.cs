using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;
using System.Windows.Threading;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.IO;
using Gryzak.Models;
using Gryzak.Services;
using static Gryzak.Services.Logger;

namespace Gryzak.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ApiService _apiService;
        private readonly ConfigService _configService;
        private readonly Dictionary<string, string> _countryMap = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _viesMap = new Dictionary<string, string>();
        private bool _isLoading;
        private bool _isLoadingMore;
        private bool _isApiConfigured;
        private string _statusFilter = "Wszystkie statusy";
        private string _searchText = "";
        private string _totalOrdersText = "0";
        private int _currentPage = 1;
        private bool _hasMorePages = true;
        private Order? _selectedOrder;
        private bool _isSubiektActive = false;
        private string _progressText = "";
        private double _progressValue = 0;
        private bool _isProgressVisible = false;
        private Views.SplashWindow? _splashWindow;
        private DispatcherTimer? _autoReleaseLicenseTimer;
        private DispatcherTimer? _countdownTimer;
        private DateTime _lastActivityTime = DateTime.Now;
        private string _timeUntilReleaseText = "";
        private bool _isReleasingLicense = false; // Flaga zapobiegająca wielokrotnemu zwolnieniu
        private List<string> _orderHistory = new List<string>();

        public ObservableCollection<Order> AllOrders { get; } = new ObservableCollection<Order>();
        public ObservableCollection<Order> FilteredOrders { get; } = new ObservableCollection<Order>();
        public ObservableCollection<string> RecentOrders { get; } = new ObservableCollection<string>();

        public Order? SelectedOrder
        {
            get => _selectedOrder;
            set
            {
                _selectedOrder = value;
                OnPropertyChanged();
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(); }
        }

        public bool IsLoadingMore
        {
            get => _isLoadingMore;
            set { _isLoadingMore = value; OnPropertyChanged(); }
        }

        public bool HasMorePages => _hasMorePages;

        public bool IsApiConfigured
        {
            get => _isApiConfigured;
            set { _isApiConfigured = value; OnPropertyChanged(); }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (_searchText != value)
                {
                    ResetActivityTimer(); // Resetuj timer przy wpisywaniu
                    _searchText = value;
                    OnPropertyChanged();
                    FilterOrders();
                }
            }
        }

        public string StatusFilter
        {
            get => _statusFilter;
            set
            {
                if (_statusFilter != value)
                {
                    ResetActivityTimer(); // Resetuj timer przy zmianie filtra
                    _statusFilter = value;
                    OnPropertyChanged();
                    FilterOrders();
                }
            }
        }

        public string TotalOrdersText
        {
            get => _totalOrdersText;
            set { _totalOrdersText = value; OnPropertyChanged(); }
        }

        public bool HasOrders => FilteredOrders.Count > 0;

        public bool IsSubiektActive
        {
            get => _isSubiektActive;
            set
            {
                _isSubiektActive = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SubiektStatusText));
                
                // Jeśli Subiekt nie jest aktywny, wyczyść countdown i resetuj flagę
                if (!value)
                {
                    TimeUntilReleaseText = "";
                    _isReleasingLicense = false;
                }
                else
                {
                    // Jeśli Subiekt staje się aktywny, upewnij się, że countdown timer działa
                    if (_countdownTimer != null && !_countdownTimer.IsEnabled)
                    {
                        _countdownTimer.Start();
                    }
                }
            }
        }

        public string TimeUntilReleaseText
        {
            get => _timeUntilReleaseText;
            private set
            {
                if (_timeUntilReleaseText != value)
                {
                    _timeUntilReleaseText = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(SubiektStatusText));
                }
            }
        }
        
        public string SubiektStatusText
        {
            get
            {
                if (!IsSubiektActive)
                {
                    return "";
                }
                
                var subiektConfig = _configService.LoadSubiektConfig();
                if (subiektConfig.AutoReleaseLicenseTimeoutMinutes > 0 && !string.IsNullOrWhiteSpace(TimeUntilReleaseText))
                {
                    return $"Subiekt GT aktywny (1 licencja) - zwolnienie za {TimeUntilReleaseText}";
                }
                
                return "Subiekt GT aktywny (1 licencja)";
            }
        }

        public string ProgressText
        {
            get => _progressText;
            set { _progressText = value; OnPropertyChanged(); }
        }

        public double ProgressValue
        {
            get => _progressValue;
            set { _progressValue = value; OnPropertyChanged(); }
        }

        public bool IsProgressVisible
        {
            get => _isProgressVisible;
            set { _isProgressVisible = value; OnPropertyChanged(); }
        }

        public ICommand RefreshCommand { get; }
        public ICommand ConfigureApiCommand { get; }
        public ICommand OpenSubiektSettingsCommand { get; }
        public ICommand OrderSelectedCommand { get; }
        public ICommand DodajZKCommand { get; }
        public ICommand NoweZKCommand { get; }
        public ICommand ZwolnijLicencjeCommand { get; }
        public ICommand ClearSearchCommand { get; }
        public ICommand OpenOrderFromHistoryCommand { get; }

        public MainViewModel()
        {
            _configService = new ConfigService();
            _apiService = new ApiService(_configService);

            RefreshCommand = new RelayCommand(async () => await LoadOrdersAsync(true));
            ConfigureApiCommand = new RelayCommand(() => OpenConfigDialog());
            OpenSubiektSettingsCommand = new RelayCommand(() => OpenSubiektSettingsDialog());
            OrderSelectedCommand = new RelayCommand<Order>(order => OnOrderSelected(order));
            DodajZKCommand = new RelayCommand(() => DodajZK());
            NoweZKCommand = new RelayCommand(() => DodajNoweZK());
            ZwolnijLicencjeCommand = new RelayCommand(() => ZwolnijLicencje(), () => IsSubiektActive);
            ClearSearchCommand = new RelayCommand(() => { SearchText = ""; });
            OpenOrderFromHistoryCommand = new RelayCommand<string>(orderId => OpenOrderFromHistory(orderId));

            // Zapisz się na event zmiany instancji Subiekta
            Services.SubiektService.InstancjaZmieniona += SubiektService_InstancjaZmieniona;
            
            // Sprawdź początkowy status
            InitializeAutoReleaseLicenseTimer();
            var subiektService = new Services.SubiektService();
            IsSubiektActive = subiektService.CzyInstancjaAktywna();
            
            // Jeśli Subiekt jest już aktywny, zresetuj timer
            if (IsSubiektActive)
            {
                ResetActivityTimer();
            }

            CheckApiConfiguration();
            
            // Upewnij się, że StatusFilter jest ustawione przed pierwszym ładowaniem
            StatusFilter = "Wszystkie statusy";
            
            LoadCountryMap();
            
            // Załaduj historię zamówień
            LoadOrderHistory();
            
            // Nie ładuj zamówień automatycznie - będzie to zrobione podczas splash screen
            // _ = LoadOrdersAsync(false);
        }

        private void LoadOrderHistory()
        {
            try
            {
                _orderHistory = _configService.LoadOrderHistory();
                RecentOrders.Clear();
                foreach (var orderId in _orderHistory)
                {
                    RecentOrders.Add(orderId);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd podczas ładowania historii zamówień");
            }
        }
        
        public void SetSplashWindow(Views.SplashWindow splashWindow)
        {
            _splashWindow = splashWindow;
        }
        
        private void LoadCountryMap()
        {
            try
            {
                var jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "kraje_iso2.json");
                if (!File.Exists(jsonPath))
                {
                    Warning($"Plik kraje_iso2.json nie został znaleziony: {jsonPath}", "MainViewModel");
                    return;
                }
                
                var jsonContent = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                var countries = JsonSerializer.Deserialize<List<CountryInfo>>(jsonContent);
                
                if (countries != null)
                {
                    foreach (var country in countries)
                    {
                        if (!string.IsNullOrWhiteSpace(country.code) && !string.IsNullOrWhiteSpace(country.name_pl))
                        {
                            _countryMap[country.code.ToUpperInvariant()] = country.name_pl;
                            
                            // Dodaj również mapę VIES jeśli dostępna
                            if (!string.IsNullOrWhiteSpace(country.code_vies))
                            {
                                _viesMap[country.code.ToUpperInvariant()] = country.code_vies;
                            }
                        }
                    }
                    Info($"Wczytano {_countryMap.Count} krajów z pliku kraje_iso2.json, z czego {_viesMap.Count} z kodami VIES", "MainViewModel");
                }
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd podczas wczytywania mapy krajów");
            }
        }
        
        private class CountryInfo
        {
            public string name_pl { get; set; } = "";
            public string name_en { get; set; } = "";
            public string code { get; set; } = "";
            public string? code_vies { get; set; }
        }
        
        public Dictionary<string, string> GetCountryMap()
        {
            return _countryMap;
        }
        
        public Dictionary<string, string> GetViesMap()
        {
            return _viesMap;
        }

        private void CheckApiConfiguration()
        {
            var config = _configService.LoadConfig();
            IsApiConfigured = !string.IsNullOrWhiteSpace(config.ApiUrl);
        }

        public async Task LoadOrdersAsync(bool reset = false)
        {
            ResetActivityTimer(); // Resetuj timer przy ładowaniu zamówień
            if (reset)
            {
                // Odznacz wszystkie zamówienia przed resetem
                foreach (var order in AllOrders)
                {
                    order.IsSelected = false;
                }
                
                AllOrders.Clear();
                SelectedOrder = null;
                _currentPage = 1;
                _hasMorePages = true;
                OnPropertyChanged(nameof(HasMorePages));
            }

            if (!_hasMorePages || _isLoadingMore)
                return;

            bool isFirstPage = _currentPage == 1;
            
            try
            {
                if (isFirstPage)
                {
                    IsLoading = true;
                    UpdateProgress("Ładowanie zamówień z API...", 30);
                }
                else
                {
                    IsLoadingMore = true;
                }

                var orders = await _apiService.LoadOrdersAsync(_currentPage);
                
                Debug("Załadowano stronę {_currentPage}: {orders.Count} zamówień", "MainViewModel");
                
                if (isFirstPage)
                {
                    UpdateProgress("Przetwarzanie danych zamówień...", 60);
                }

                // Jeśli API zwraca mniej niż 20 zamówień lub 0, uznaj to za ostatnią stronę
                if (orders.Count < 20)
                {
                    _hasMorePages = false;
                    OnPropertyChanged(nameof(HasMorePages));
                    Debug("Ostatnia strona - zwrócono {orders.Count} zamówień", "MainViewModel");
                }

                if (orders.Count == 0)
                {
                    Debug("Brak zamówień na stronie {_currentPage}", "MainViewModel");
                }
                else
                {
                    foreach (var order in orders)
                    {
                        // Sprawdź czy zamówienie nie jest już w liście (duplikaty)
                        if (!AllOrders.Any(o => o.Id == order.Id))
                        {
                            // Mapuj kraj z ISO code 2 na polską nazwę jeśli dostępne
                            if (!string.IsNullOrWhiteSpace(order.IsoCode2))
                            {
                                var isoCode2 = order.IsoCode2.ToUpperInvariant();
                                if (_countryMap.TryGetValue(isoCode2, out var polishName))
                                {
                                    order.Country = polishName;
                                    Debug("Zmapowano kraj (ISO {isoCode2}): {order.Country}", "MainViewModel");
                                }
                            }
                            
                            AllOrders.Add(order);
                            Debug("Dodano zamówienie: {order.Id} - {order.Customer}", "MainViewModel");
                            
                            // Jeśli to zaznaczone zamówienie, ustaw flagę
                            if (SelectedOrder != null && SelectedOrder.Id == order.Id)
                            {
                                order.IsSelected = true;
                            }
                        }
                    }

                    _currentPage++;
                }

                // Sprawdź istnienie dokumentów ZK w Subiekcie GT
                if (AllOrders.Count > 0)
                {
                    try
                    {
                        if (isFirstPage)
                        {
                            UpdateProgress("Sprawdzanie dokumentów ZK w Subiekcie GT...", 70);
                        }
                        
                        var subiektService = new Services.SubiektService();
                        var numeryZamowien = AllOrders.Select(o => o.Id).Where(id => !string.IsNullOrWhiteSpace(id)).ToList();
                        
                        if (numeryZamowien.Count > 0)
                        {
                            Debug($"Sprawdzanie istnienia {numeryZamowien.Count} dokumentów ZK w Subiekcie GT...", "MainViewModel");
                            
                            var istniejaceDokumenty = await Task.Run(() => subiektService.SprawdzIstnienieDokumentowZK(numeryZamowien));
                            
                            // Aktualizuj zamówienia na podstawie wyników SQL
                            Application.Current?.Dispatcher.Invoke(() =>
                            {
                                int zaktualizowane = 0;
                                foreach (var order in AllOrders)
                                {
                                    if (!string.IsNullOrWhiteSpace(order.Id))
                                    {
                                        var orderIdTrimmed = order.Id.Trim();
                                        if (istniejaceDokumenty.TryGetValue(orderIdTrimmed, out var numerDokumentu))
                                        {
                                            order.IsDocumentExists = true;
                                            order.SubiektDocumentNumber = numerDokumentu ?? "";
                                            zaktualizowane++;
                                            Debug($"Ustawiono IsDocumentExists=true i SubiektDocumentNumber='{numerDokumentu}' dla zamówienia: {order.Id}", "MainViewModel");
                                        }
                                        else
                                        {
                                            // Resetuj jeśli dokument nie istnieje
                                            order.IsDocumentExists = false;
                                            order.SubiektDocumentNumber = "";
                                        }
                                    }
                                }
                                Debug($"Zaktualizowano flagę IsDocumentExists dla {zaktualizowane} zamówień z {istniejaceDokumenty.Count} znalezionych dokumentów.", "MainViewModel");
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "MainViewModel", "Błąd podczas sprawdzania istnienia dokumentów ZK");
                        // Nie blokuj dalszego przetwarzania w przypadku błędu
                    }
                }

                FilterOrders();
                UpdateStatistics();
                OnPropertyChanged(nameof(HasOrders));
                
                if (isFirstPage)
                {
                    UpdateProgress("Zamówienia załadowane", 85);
                    
                    // Jeśli ładowaliśmy podczas splash screena, ukryj progress bar
                    // (splash screen zostanie zamknięty w App.xaml.cs)
                    if (_splashWindow != null)
                    {
                        await Task.Delay(100); // Krótkie opóźnienie aby użytkownik zobaczył "Zamówienia załadowane"
                        HideProgress();
                    }
                    else
                    {
                        // Normalne ładowanie (bez splash screen) - ukryj po chwili
                        await Task.Delay(500);
                        HideProgress();
                    }
                }
                
                Debug("Przetworzonych zamówień: {AllOrders.Count}, przefiltrowanych: {FilteredOrders.Count}, HasOrders: {HasOrders}, HasMorePages: {_hasMorePages}", "MainViewModel");
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd ładowania");
                if (isFirstPage)
                {
                    UpdateProgress("Błąd ładowania", 0);
                    await Task.Delay(1000);
                    HideProgress();
                    
                    System.Windows.MessageBox.Show(
                        $"Błąd ładowania zamówień: {ex.Message}",
                        "Błąd",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
                // Przy błędzie na kolejnych stronach nie pokazujemy komunikatu użytkownikowi
            }
            finally
            {
                IsLoading = false;
                IsLoadingMore = false;
                Debug("IsLoading: {IsLoading}, IsLoadingMore: {IsLoadingMore}", "MainViewModel");
            }
        }

        public async Task LoadNextPageAsync()
        {
            if (!_hasMorePages || _isLoadingMore || _isLoading)
                return;

            await LoadOrdersAsync(false);
        }

        private void FilterOrders()
        {
            FilteredOrders.Clear();
            
            Debug("FilterOrders - StatusFilter: '{StatusFilter}', SearchText: '{SearchText}', AllOrders.Count: {AllOrders.Count}", "MainViewModel");
            
            if (AllOrders.Count > 0)
            {
                Debug("Przykładowy status pierwszego zamówienia: '{AllOrders[0].Status}'", "MainViewModel");
            }

            // Najpierw filtruj po statusie
            var orders = StatusFilter == "Wszystkie statusy"
                ? AllOrders
                : AllOrders.Where(o => o.Status == StatusFilter);

            // Potem filtruj po tekście wyszukiwania (ID, klient, email, telefon, NIP, kraj, kod ISO)
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var searchLower = SearchText.ToLower();
                orders = orders.Where(o =>
                    o.Id.ToLower().Contains(searchLower) ||
                    (o.Customer?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Email?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Phone?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Nip?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Company?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Address?.ToLower().Contains(searchLower) ?? false) ||
                    (o.Country?.ToLower().Contains(searchLower) ?? false) ||
                    (o.IsoCode3?.ToLower().Contains(searchLower) ?? false) ||
                    (o.CountryWithIso3?.ToLower().Contains(searchLower) ?? false)
                );
            }

            var ordersList = orders.ToList();
            Debug("Po filtrowaniu: {ordersList.Count} zamówień", "MainViewModel");

            foreach (var order in ordersList)
            {
                FilteredOrders.Add(order);
            }

            Debug("FilteredOrders.Count: {FilteredOrders.Count}", "MainViewModel");
            OnPropertyChanged(nameof(HasOrders));
            UpdateStatistics();
        }

        private void UpdateStatistics()
        {
            var orders = FilteredOrders;
            TotalOrdersText = orders.Count.ToString();
        }

        public void UpdateProgress(string text, double value)
        {
            ProgressText = text;
            ProgressValue = value;
            IsProgressVisible = !string.IsNullOrEmpty(text);
            
            // Aktualizuj splash screen jeśli jest dostępny
            if (_splashWindow != null)
            {
                _splashWindow.UpdateProgress(value, text);
            }
        }

        public void HideProgress()
        {
            IsProgressVisible = false;
            ProgressText = "";
            ProgressValue = 0;
        }

        private void OpenConfigDialog()
        {
            var configWindow = new Views.ConfigDialog(_configService);
            configWindow.ShowDialog();
            
            // Sprawdź konfigurację i odśwież listę
            CheckApiConfiguration();
            _ = LoadOrdersAsync();
        }

        private void OpenSubiektSettingsDialog()
        {
            var settingsWindow = new Views.SubiektSettingsDialog(_configService);
            settingsWindow.ShowDialog();
        }

        private void DodajZK()
        {
            ResetActivityTimer(); // Resetuj timer przy aktywności
            
            // Dodaj zamówienie do historii (na początku metody, przed wywołaniem SubiektService)
            if (!string.IsNullOrWhiteSpace(SelectedOrder?.Id))
            {
                _configService.AddOrderToHistory(SelectedOrder.Id);
                UpdateRecentOrders();
            }
            
            // Upewnij się że wartości kosztów są brutto (konwertuj z netto jeśli potrzeba)
            EnsureCostsAreBrutto(SelectedOrder);
            
            // Pobierz NIP z wybranego zamówienia
            var nip = SelectedOrder?.Nip;
            
            // Nie pokazujemy już MessageBox - dialog wyboru kontrahenta obsługuje wybór
            var subiektService = new Services.SubiektService();
            var orderEmail = SelectedOrder?.Email;
            // Jeśli email to "Brak email", traktuj jako null
            if (orderEmail == "Brak email" || string.IsNullOrWhiteSpace(orderEmail))
            {
                orderEmail = null;
            }
            subiektService.OtworzOknoZK(nip, SelectedOrder?.Items, SelectedOrder?.CouponAmount, SelectedOrder?.SubTotal, SelectedOrder?.CouponTitle, SelectedOrder?.Id, SelectedOrder?.HandlingAmountNetto, SelectedOrder?.ShippingAmountNetto, SelectedOrder?.Currency, SelectedOrder?.CurrencyValue, SelectedOrder?.CodFeeAmountNetto, SelectedOrder?.Total, SelectedOrder?.GlsAmountNetto, SelectedOrder?.GlsKgAmountNetto, orderEmail, SelectedOrder?.Customer, SelectedOrder?.Phone, SelectedOrder?.Company, SelectedOrder?.Address, SelectedOrder?.PaymentAddress1, SelectedOrder?.PaymentAddress2, SelectedOrder?.PaymentPostcode, SelectedOrder?.PaymentCity, SelectedOrder?.Country, SelectedOrder?.IsoCode2, SelectedOrder?.UseEuVatRate ?? false);
            // Status zostanie zaktualizowany przez event InstancjaZmieniona
        }
        
        private void EnsureCostsAreBrutto(Order? order)
        {
            if (order == null) return;
            
            // UWAGA: Wartości kosztów są już konwertowane z netto na brutto w UpdateOrderFromDetails
            // (linie 1664-1704), więc nie ma potrzeby ponownej konwersji tutaj.
            // Funkcja pozostaje dla kompatybilności wstecznej, ale nie wykonuje żadnych operacji.
            // Wartości z totals API są zawsze netto i są konwertowane na brutto (VAT 23%) w UpdateOrderFromDetails.
        }

        private void DodajNoweZK()
        {
            ResetActivityTimer(); // Resetuj timer przy aktywności
            
            // Dodaj zamówienie do historii (na początku metody, przed wywołaniem SubiektService)
            if (!string.IsNullOrWhiteSpace(SelectedOrder?.Id))
            {
                _configService.AddOrderToHistory(SelectedOrder.Id);
                UpdateRecentOrders();
            }
            
            // Upewnij się że wartości kosztów są brutto (konwertuj z netto jeśli potrzeba)
            EnsureCostsAreBrutto(SelectedOrder);
            
            // Pobierz NIP z wybranego zamówienia
            var nip = SelectedOrder?.Nip;
            
            var subiektService = new Services.SubiektService();
            var orderEmail = SelectedOrder?.Email;
            // Jeśli email to "Brak email", traktuj jako null
            if (orderEmail == "Brak email" || string.IsNullOrWhiteSpace(orderEmail))
            {
                orderEmail = null;
            }
            // Użyj OtworzNoweZK - zawsze tworzy nowy dokument bez sprawdzania istnienia
            subiektService.OtworzNoweZK(nip, SelectedOrder?.Items, SelectedOrder?.CouponAmount, SelectedOrder?.SubTotal, SelectedOrder?.CouponTitle, SelectedOrder?.Id, SelectedOrder?.HandlingAmountNetto, SelectedOrder?.ShippingAmountNetto, SelectedOrder?.Currency, SelectedOrder?.CurrencyValue, SelectedOrder?.CodFeeAmountNetto, SelectedOrder?.Total, SelectedOrder?.GlsAmountNetto, SelectedOrder?.GlsKgAmountNetto, orderEmail, SelectedOrder?.Customer, SelectedOrder?.Phone, SelectedOrder?.Company, SelectedOrder?.Address, SelectedOrder?.PaymentAddress1, SelectedOrder?.PaymentAddress2, SelectedOrder?.PaymentPostcode, SelectedOrder?.PaymentCity, SelectedOrder?.Country, SelectedOrder?.IsoCode2, SelectedOrder?.UseEuVatRate ?? false);
            // Status zostanie zaktualizowany przez event InstancjaZmieniona
        }

        private void UpdateRecentOrders()
        {
            try
            {
                _orderHistory = _configService.LoadOrderHistory();
                RecentOrders.Clear();
                foreach (var orderId in _orderHistory)
                {
                    RecentOrders.Add(orderId);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd podczas aktualizacji listy ostatnich zamówień");
            }
        }

        private async void OpenOrderFromHistory(string? orderId)
        {
            Debug($"OpenOrderFromHistory wywołana z orderId: {orderId}", "MainViewModel");
            
            if (string.IsNullOrWhiteSpace(orderId))
            {
                Debug("OpenOrderFromHistory: orderId jest pusty, przerywam", "MainViewModel");
                return;
            }

            try
            {
                Debug($"OpenOrderFromHistory: Próbuję otworzyć zamówienie {orderId}", "MainViewModel");
                
                // Upewnij się, że jesteśmy na wątku UI
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    Debug("OpenOrderFromHistory: Jesteśmy na wątku UI, wywołuję OpenOrderFromHistoryInternal", "MainViewModel");
                    await OpenOrderFromHistoryInternal(orderId);
                }
                else
                {
                    Debug("OpenOrderFromHistory: Nie jesteśmy na wątku UI, wywołuję przez Dispatcher", "MainViewModel");
                    await Application.Current.Dispatcher.InvokeAsync(async () => await OpenOrderFromHistoryInternal(orderId));
                }
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", $"Błąd podczas otwierania zamówienia z historii: {orderId}");
                MessageBox.Show(
                    $"Błąd podczas otwierania zamówienia: {ex.Message}",
                    "Błąd",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task OpenOrderFromHistoryInternal(string orderId)
        {
            Debug($"OpenOrderFromHistoryInternal: Szukam zamówienia {orderId}", "MainViewModel");
            
            // Najpierw spróbuj znaleźć zamówienie w aktualnej liście
            var order = AllOrders.FirstOrDefault(o => o.Id == orderId);
            
            if (order != null)
            {
                Debug($"OpenOrderFromHistoryInternal: Znaleziono zamówienie {orderId} w liście, zaznaczam i otwieram okno szczegółów", "MainViewModel");
                
                // Zaznacz zamówienie (bez wywoływania OnOrderSelected, żeby nie blokować)
                SelectedOrder = order;
                
                // Otwórz okno szczegółów bezpośrednio (podobnie jak przy podwójnym kliknięciu)
                // Znajdź MainWindow w otwartych oknach aplikacji
                Views.MainWindow? mainWindow = null;
                foreach (Window window in Application.Current.Windows)
                {
                    if (window is Views.MainWindow mw)
                    {
                        mainWindow = mw;
                        break;
                    }
                }
                
                Debug($"OpenOrderFromHistoryInternal: mainWindow = {(mainWindow != null ? "nie null" : "null")}", "MainViewModel");
                
                if (mainWindow != null)
                {
                    Debug($"OpenOrderFromHistoryInternal: Wywołuję OpenOrderDetailsDialog dla zamówienia {orderId}", "MainViewModel");
                    mainWindow.OpenOrderDetailsDialog(order, this);
                    Debug($"OpenOrderFromHistoryInternal: OpenOrderDetailsDialog wywołane", "MainViewModel");
                }
                else
                {
                    // Fallback: otwórz okno bezpośrednio bez Owner
                    Debug($"OpenOrderFromHistoryInternal: MainWindow nie znaleziony, otwieram OrderDetailsDialog bezpośrednio", "MainViewModel");
                    var detailsDialog = new Views.OrderDetailsDialog(order, this);
                    detailsDialog.ShowDialog();
                }
            }
            else
            {
                Debug($"OpenOrderFromHistoryInternal: Nie znaleziono zamówienia {orderId} w liście, pobieram z API", "MainViewModel");
                
                // Jeśli nie znaleziono w liście, pobierz zamówienie bezpośrednio z API
                try
                {
                    Debug($"OpenOrderFromHistoryInternal: Pobieram szczegóły zamówienia {orderId} z API", "MainViewModel");
                    var detailsJson = await _apiService.GetOrderDetailsAsync(orderId);
                    
                    Debug($"OpenOrderFromHistoryInternal: Utworzenie obiektu Order z JSON dla zamówienia {orderId}", "MainViewModel");
                    order = CreateOrderFromDetailsJson(detailsJson);
                    
                    if (order != null)
                    {
                        Debug($"OpenOrderFromHistoryInternal: Utworzono zamówienie {orderId}, otwieram okno szczegółów", "MainViewModel");
                        // Zaznacz zamówienie (nie dodajemy do AllOrders - tylko wyświetlamy w oknie szczegółów)
                        SelectedOrder = order;
                        
                        // Znajdź MainWindow w otwartych oknach aplikacji
                        Views.MainWindow? mainWindow = null;
                        foreach (Window window in Application.Current.Windows)
                        {
                            if (window is Views.MainWindow mw)
                            {
                                mainWindow = mw;
                                break;
                            }
                        }
                        
                        if (mainWindow != null)
                        {
                            mainWindow.OpenOrderDetailsDialog(order, this);
                        }
                        else
                        {
                            // Fallback: otwórz okno bezpośrednio bez Owner
                            var detailsDialog = new Views.OrderDetailsDialog(order, this);
                            detailsDialog.ShowDialog();
                        }
                    }
                    else
                    {
                        Error($"Nie udało się utworzyć obiektu Order z JSON dla zamówienia {orderId}", "MainViewModel");
                        MessageBox.Show(
                            $"Nie udało się załadować zamówienia o numerze: {orderId}",
                            "Błąd",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    Error(ex, "MainViewModel", $"Błąd podczas pobierania zamówienia {orderId} z API");
                    MessageBox.Show(
                        $"Błąd podczas pobierania zamówienia: {ex.Message}",
                        "Błąd",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private void SubiektService_InstancjaZmieniona(object? sender, bool aktywna)
        {
            // Aktualizuj status w wątku UI
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                IsSubiektActive = aktywna;
                // Powiadom o zmianie dostępności komendy
                if (ZwolnijLicencjeCommand is RelayCommand relayCommand)
                {
                    relayCommand.RaiseCanExecuteChanged();
                }
                
                // Jeśli Subiekt jest aktywny, zrestartuj timer automatycznego zwalniania
                if (aktywna)
                {
                    ResetActivityTimer();
                    _isReleasingLicense = false; // Resetuj flagę gdy Subiekt staje się aktywny
                    
                    // Upewnij się, że countdown timer jest uruchomiony
                    if (_countdownTimer != null && !_countdownTimer.IsEnabled)
                    {
                        _countdownTimer.Start();
                        Debug("Countdown timer został uruchomiony ponownie po aktywacji Subiekta.", "MainViewModel");
                    }
                }
                else
                {
                    StopAutoReleaseLicenseTimer();
                    _isReleasingLicense = false; // Resetuj flagę gdy Subiekt nie jest aktywny
                }
            });
        }
        
        /// <summary>
        /// Inicjalizuje timer automatycznego zwalniania licencji
        /// </summary>
        private void InitializeAutoReleaseLicenseTimer()
        {
            _autoReleaseLicenseTimer = new DispatcherTimer();
            _autoReleaseLicenseTimer.Tick += AutoReleaseLicenseTimer_Tick;
            _autoReleaseLicenseTimer.Interval = TimeSpan.FromMinutes(1); // Sprawdzaj co minutę
            _autoReleaseLicenseTimer.Start();
            _lastActivityTime = DateTime.Now;
            
            // Timer do wyświetlania countdown
            _countdownTimer = new DispatcherTimer();
            _countdownTimer.Tick += CountdownTimer_Tick;
            _countdownTimer.Interval = TimeSpan.FromSeconds(1); // Aktualizuj co sekundę
            _countdownTimer.Start();
            
            Debug("Timer automatycznego zwalniania licencji został uruchomiony.", "MainViewModel");
        }
        
        /// <summary>
        /// Obsługa zdarzenia Tick timera countdown - aktualizuje wyświetlany czas do zwolnienia
        /// </summary>
        private void CountdownTimer_Tick(object? sender, EventArgs e)
        {
            // Aktualizuj tylko jeśli Subiekt jest aktywny
            if (!IsSubiektActive)
            {
                TimeUntilReleaseText = "";
                return;
            }
            
            // Wczytaj konfigurację
            var subiektConfig = _configService.LoadSubiektConfig();
            
            // Jeśli timeout jest 0, wyłączone - nie pokazuj countdown
            if (subiektConfig.AutoReleaseLicenseTimeoutMinutes <= 0)
            {
                TimeUntilReleaseText = "";
                return;
            }
            
            // Oblicz czas nieaktywności
            var inactiveTime = DateTime.Now - _lastActivityTime;
            var inactiveMinutes = inactiveTime.TotalMinutes;
            
            // Oblicz pozostały czas
            var remainingMinutes = subiektConfig.AutoReleaseLicenseTimeoutMinutes - inactiveMinutes;
            
            if (remainingMinutes <= 0)
            {
                TimeUntilReleaseText = "0:00";
                
                // Zwolnij licencję natychmiast gdy czas dobił do 0 (tylko raz)
                if (!_isReleasingLicense && IsSubiektActive)
                {
                    _isReleasingLicense = true;
                    
                    Debug($"Czas nieaktywności ({inactiveMinutes:F1} minut) przekroczył limit ({subiektConfig.AutoReleaseLicenseTimeoutMinutes} minut). Zwolnienie licencji...", "MainViewModel");
                    
                    try
                    {
                        var subiektService = new Services.SubiektService();
                        subiektService.ZwolnijLicencje();
                        
                        // Wyświetl MessageBox
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(
                                $"Licencja Subiekta GT została automatycznie zwolniona po {subiektConfig.AutoReleaseLicenseTimeoutMinutes} minutach nieaktywności.",
                                "Licencja zwolniona",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                        });
                        
                        Info($"Licencja Subiekta GT została automatycznie zwolniona po {subiektConfig.AutoReleaseLicenseTimeoutMinutes} minutach nieaktywności.", "MainViewModel");
                    }
                    catch (Exception ex)
                    {
                        Error(ex, "MainViewModel", "Błąd podczas automatycznego zwalniania licencji");
                        _isReleasingLicense = false; // Resetuj flagę w przypadku błędu
                    }
                }
            }
            else
            {
                var minutes = (int)remainingMinutes;
                var seconds = (int)((remainingMinutes - minutes) * 60);
                TimeUntilReleaseText = $"{minutes}:{seconds:D2}";
            }
        }
        
        /// <summary>
        /// Zatrzymuje timer automatycznego zwalniania licencji
        /// </summary>
        private void StopAutoReleaseLicenseTimer()
        {
            if (_autoReleaseLicenseTimer != null)
            {
                _autoReleaseLicenseTimer.Stop();
                Debug("Timer automatycznego zwalniania licencji został zatrzymany.", "MainViewModel");
            }
            
            if (_countdownTimer != null)
            {
                _countdownTimer.Stop();
            }
            
            TimeUntilReleaseText = "";
        }
        
        /// <summary>
        /// Resetuje timer aktywności - wywoływane przy każdej aktywności użytkownika
        /// </summary>
        public void ResetActivityTimer()
        {
            _lastActivityTime = DateTime.Now;
            // Resetuj flagę zwalniania, aby licencja mogła być zwolniona ponownie po następnej aktywności
            _isReleasingLicense = false;
            // Nie loguj każdej aktywności, żeby nie spamować logów
        }
        
        /// <summary>
        /// Obsługa zdarzenia Tick timera - sprawdza czy należy zwolnić licencję
        /// </summary>
        private void AutoReleaseLicenseTimer_Tick(object? sender, EventArgs e)
        {
            // Sprawdź tylko jeśli Subiekt jest aktywny
            if (!IsSubiektActive)
            {
                return;
            }
            
            // Wczytaj konfigurację
            var subiektConfig = _configService.LoadSubiektConfig();
            
            // Jeśli timeout jest 0, wyłączone
            if (subiektConfig.AutoReleaseLicenseTimeoutMinutes <= 0)
            {
                return;
            }
            
            // Oblicz czas nieaktywności
            var inactiveTime = DateTime.Now - _lastActivityTime;
            var inactiveMinutes = inactiveTime.TotalMinutes;
            
            // Jeśli czas nieaktywności przekroczył limit
            if (inactiveMinutes >= subiektConfig.AutoReleaseLicenseTimeoutMinutes)
            {
                Debug($"Czas nieaktywności ({inactiveMinutes:F1} minut) przekroczył limit ({subiektConfig.AutoReleaseLicenseTimeoutMinutes} minut). Zwolnienie licencji...", "MainViewModel");
                
                // Zwolnij licencję
                try
                {
                    var subiektService = new Services.SubiektService();
                    subiektService.ZwolnijLicencje();
                    
                    // Wyświetl MessageBox
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(
                            $"Licencja Subiekta GT została automatycznie zwolniona po {subiektConfig.AutoReleaseLicenseTimeoutMinutes} minutach nieaktywności.",
                            "Licencja zwolniona",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    });
                    
                    Info($"Licencja Subiekta GT została automatycznie zwolniona po {subiektConfig.AutoReleaseLicenseTimeoutMinutes} minutach nieaktywności.", "MainViewModel");
                }
                catch (Exception ex)
                {
                    Error(ex, "MainViewModel", "Błąd podczas automatycznego zwalniania licencji");
                }
            }
        }

        private void ZwolnijLicencje()
        {
            var subiektService = new Services.SubiektService();
            subiektService.ZwolnijLicencje();
            // Status zostanie zaktualizowany przez event InstancjaZmieniona
            System.Windows.MessageBox.Show(
                "Licencja Subiekta GT została zwolniona.\n\nInstancja została zamknięta.",
                "Informacja",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        private void OnOrderSelected(Order? order)
        {
            ResetActivityTimer(); // Resetuj timer przy aktywności
            // Odznacz poprzednio zaznaczone zamówienie
            if (SelectedOrder != null)
            {
                SelectedOrder.IsSelected = false;
            }

            SelectedOrder = order;

            // Zaznacz nowe zamówienie
            if (order != null)
            {
                order.IsSelected = true;
                
                // Wypisz pełne dane zamówienia do konsoli
                Debug("=== ZAZNACZONO ZAMÓWIENIE ===", "MainViewModel");
                Debug("ID Zamówienia: {order.Id}", "MainViewModel");
                Debug("Klient: {order.Customer}", "MainViewModel");
                Debug("Email: {order.Email}", "MainViewModel");
                Debug("Telefon: {order.Phone}", "MainViewModel");
                Debug($"Firma: {order.Company ?? "Brak"}", "MainViewModel");
                Debug($"NIP: {order.Nip ?? "Brak"}", "MainViewModel");
                Debug($"Adres: {order.Address ?? "Brak"}", "MainViewModel");
                Debug("Status: {order.Status}", "MainViewModel");
                Debug("Status Płatności: {order.PaymentStatus}", "MainViewModel");
                Debug("Wartość: {order.Total} {order.Currency}", "MainViewModel");
                Debug("Data: {order.Date:dd.MM.yyyy HH:mm}", "MainViewModel");
                Debug("Obsługuje: {order.AssignedTo}", "MainViewModel");
                Debug("Liczba produktów: {order.Items?.Count ?? 0}", "MainViewModel");
                Debug("==========================", "MainViewModel");
                
                // Pobierz szczegóły zamówienia z API
                LoadOrderDetailsAsync(order.Id);
            }
        }
        
        private async void LoadOrderDetailsAsync(string orderId)
        {
            try
            {
                Debug("Ładowanie szczegółów zamówienia {orderId}...", "MainViewModel");
                var detailsJson = await _apiService.GetOrderDetailsAsync(orderId);
                
                // Sformatuj JSON do pretty print
                string formattedJson;
                try
                {
                    using var doc = JsonDocument.Parse(detailsJson);
                    formattedJson = JsonSerializer.Serialize(doc.RootElement, new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    });
                }
                catch
                {
                    // Jeśli nie można sparsować jako JSON, użyj oryginalnego tekstu
                    formattedJson = detailsJson;
                }
                
                Debug("=== SZCZEGÓŁY ZAMÓWIENIA Z API ===", "MainViewModel");
                Debug(formattedJson, "MainViewModel");
                Debug("===================================", "MainViewModel");
                
                // Zaktualizuj zamówienie w liście
                UpdateOrderFromDetails(SelectedOrder, detailsJson);
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd ładowania szczegółów zamówienia");
            }
        }
        
        private Order? CreateOrderFromDetailsJson(string detailsJson)
        {
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(detailsJson);
                var root = doc.RootElement;
                
                // Podstawowe informacje o zamówieniu
                var orderId = root.TryGetProperty("order_id", out var id) ? id.GetString() ?? "" : "";
                if (string.IsNullOrWhiteSpace(orderId))
                {
                    Error("Nie znaleziono order_id w JSON szczegółów zamówienia", "MainViewModel");
                    return null;
                }
                
                var firstname = root.TryGetProperty("firstname", out var fn) ? fn.GetString() ?? "" : "";
                var lastname = root.TryGetProperty("lastname", out var ln) ? ln.GetString() ?? "" : "";
                var email = root.TryGetProperty("email", out var e) ? e.GetString() ?? "" : "";
                var telephone = root.TryGetProperty("telephone", out var tel) ? tel.GetString() ?? "" : "";
                var company = root.TryGetProperty("payment_company", out var comp) ? comp.GetString() : null;
                var vat = root.TryGetProperty("vat", out var v) ? v.GetString() : null;
                var address1 = root.TryGetProperty("payment_address_1", out var a1) ? a1.GetString() ?? "" : "";
                var address2 = root.TryGetProperty("payment_address_2", out var a2) ? a2.GetString() : null;
                var postcode = root.TryGetProperty("payment_postcode", out var pc) ? pc.GetString() ?? "" : "";
                var city = root.TryGetProperty("payment_city", out var c) ? c.GetString() ?? "" : "";
                var country = root.TryGetProperty("payment_country", out var countryProp) ? countryProp.GetString() ?? "" : "";
                var status = root.TryGetProperty("status", out var s) ? s.GetString() ?? "" : "";
                var paymentStatus = root.TryGetProperty("payment_status", out var ps) ? ps.GetString() ?? "Nieznany" : "Nieznany";
                
                // Obsługa total jako liczby float/double lub stringa
                string total = "0";
                if (root.TryGetProperty("total", out var t))
                {
                    if (t.ValueKind == System.Text.Json.JsonValueKind.Number)
                    {
                        var totalValue = t.GetDouble();
                        total = totalValue.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else if (t.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        total = t.GetString() ?? "0";
                    }
                }
                
                var currency = root.TryGetProperty("currency_code", out var curr) ? curr.GetString() ?? "PLN" : "PLN";
                double? currencyValue = null;
                if (root.TryGetProperty("currency_value", out var currValue))
                {
                    if (currValue.ValueKind == System.Text.Json.JsonValueKind.Number)
                    {
                        currencyValue = currValue.GetDouble();
                    }
                    else if (currValue.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var currValueStr = currValue.GetString();
                        if (double.TryParse(currValueStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            currencyValue = parsedValue;
                        }
                    }
                }
                var dateAdded = root.TryGetProperty("date_added", out var date) ? date.GetString() ?? "" : "";
                
                // Parsuj datę
                DateTime dateParsed = DateTime.Now;
                if (!string.IsNullOrWhiteSpace(dateAdded))
                {
                    if (DateTime.TryParse(dateAdded, out var parsed))
                    {
                        dateParsed = parsed;
                    }
                }
                
                // Buduj adres
                var addressParts = new List<string>();
                if (!string.IsNullOrWhiteSpace(address1))
                {
                    var addr1 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address1));
                    addressParts.Add(addr1);
                    if (!string.IsNullOrWhiteSpace(address2))
                    {
                        var addr2 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address2));
                        addressParts.Add(addr2);
                    }
                    var decodedPostcode = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(postcode));
                    var decodedCity = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(city));
                    if (!string.IsNullOrWhiteSpace(decodedPostcode) || !string.IsNullOrWhiteSpace(decodedCity))
                    {
                        addressParts.Add($"{decodedPostcode} {decodedCity}".Trim());
                    }
                }
                var address = string.Join(", ", addressParts.Where(s => !string.IsNullOrWhiteSpace(s)));
                
                // Dekoduj encje HTML (czasem podwójnie zakodowane)
                if (!string.IsNullOrWhiteSpace(company))
                {
                    company = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(company));
                }
                
                // ISO codes
                var isoCode2 = root.TryGetProperty("payment_iso_code_2", out var iso2) ? iso2.GetString() : 
                               (root.TryGetProperty("iso_code_2", out var iso2Fallback) ? iso2Fallback.GetString() : null);
                var isoCode3 = root.TryGetProperty("iso_code_3", out var iso3) ? iso3.GetString() : null;
                
                // Mapuj kraj używając ISO code
                string mappedCountry = country;
                if (!string.IsNullOrWhiteSpace(isoCode2))
                {
                    var iso2Upper = isoCode2.ToUpperInvariant();
                    if (_countryMap.TryGetValue(iso2Upper, out var polishName))
                    {
                        mappedCountry = polishName;
                    }
                }
                
                // Sformatuj total
                string formattedTotal = "0.00";
                if (!string.IsNullOrWhiteSpace(total) && total != "0")
                {
                    if (double.TryParse(total, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out var totalVal))
                    {
                        formattedTotal = totalVal.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        formattedTotal = total;
                    }
                }
                
                // Utwórz obiekt Order
                var order = new Order
                {
                    Id = orderId,
                    Customer = $"{firstname} {lastname}".Trim(),
                    Email = string.IsNullOrWhiteSpace(email) ? "Brak email" : email,
                    Phone = string.IsNullOrWhiteSpace(telephone) ? "Brak telefonu" : telephone,
                    Company = company,
                    Nip = vat,
                    Address = string.IsNullOrWhiteSpace(address) ? null : address,
                    PaymentAddress1 = string.IsNullOrWhiteSpace(address1) ? null : System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address1)),
                    PaymentAddress2 = string.IsNullOrWhiteSpace(address2) ? null : System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address2)),
                    PaymentPostcode = string.IsNullOrWhiteSpace(postcode) ? null : System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(postcode)),
                    PaymentCity = string.IsNullOrWhiteSpace(city) ? null : System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(city)),
                    Status = status,
                    PaymentStatus = paymentStatus,
                    Total = formattedTotal,
                    Currency = currency,
                    CurrencyValue = currencyValue,
                    Date = dateParsed,
                    Country = mappedCountry,
                    IsoCode2 = isoCode2,
                    IsoCode3 = isoCode3
                };
                
                // Teraz użyj UpdateOrderFromDetails aby zaktualizować produkty i totals (używa tej samej logiki)
                UpdateOrderFromDetails(order, detailsJson);
                
                Debug($"Utworzono zamówienie {orderId} z szczegółów API", "MainViewModel");
                return order;
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd podczas tworzenia Order z JSON szczegółów");
                return null;
            }
        }
        
        private void UpdateOrderFromDetails(Order? order, string detailsJson)
        {
            if (order == null) return;
            
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(detailsJson);
                var root = doc.RootElement;
                
                // Aktualizuj email
                if (root.TryGetProperty("email", out var emailProp) && emailProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Email = emailProp.GetString() ?? "Brak email";
                    Debug("Zaktualizowano email: {order.Email}", "MainViewModel");
                }
                
                // Aktualizuj telefon
                if (root.TryGetProperty("telephone", out var telProp) && telProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Phone = telProp.GetString() ?? "Brak telefonu";
                    Debug("Zaktualizowano telefon: {order.Phone}", "MainViewModel");
                }
                
                // Aktualizuj nazwę firmy (payment_company)
                if (root.TryGetProperty("payment_company", out var companyProp) && companyProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    var companyValue = companyProp.GetString();
                    if (!string.IsNullOrWhiteSpace(companyValue))
                    {
                        // Decoduj HTML entities dwukrotnie (bo API zwraca podwójnie zakodowane encje)
                        companyValue = System.Net.WebUtility.HtmlDecode(companyValue);
                        companyValue = System.Net.WebUtility.HtmlDecode(companyValue);
                        order.Company = companyValue;
                        Debug("Zaktualizowano firmę: {order.Company}", "MainViewModel");
                    }
                }
                
                // Aktualizuj adres
                if (root.TryGetProperty("payment_address_1", out var addr1Prop) && addr1Prop.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    var addressParts = new List<string>();
                    var addr1Raw = addr1Prop.GetString() ?? "";
                    var addr1 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(addr1Raw));
                    addressParts.Add(addr1);
                    
                    // Dodaj address_2 jeśli istnieje
                    if (root.TryGetProperty("payment_address_2", out var addr2Prop) && addr2Prop.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var addr2Raw = addr2Prop.GetString();
                        if (!string.IsNullOrWhiteSpace(addr2Raw))
                        {
                            var addr2 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(addr2Raw));
                            addressParts.Add(addr2);
                        }
                    }
                    
                    // Dodaj postcode i city
                    var postcode = "";
                    var city = "";
                    if (root.TryGetProperty("payment_postcode", out var pcProp) && pcProp.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var pcRaw = pcProp.GetString() ?? "";
                        postcode = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(pcRaw));
                    }
                    if (root.TryGetProperty("payment_city", out var cityProp) && cityProp.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var cityRaw = cityProp.GetString() ?? "";
                        city = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(cityRaw));
                    }
                    
                    if (!string.IsNullOrWhiteSpace(postcode) || !string.IsNullOrWhiteSpace(city))
                    {
                        addressParts.Add($"{postcode} {city}".Trim());
                    }
                    
                    order.Address = string.Join(", ", addressParts.Where(s => !string.IsNullOrWhiteSpace(s)));
                    Debug("Zaktualizowano adres: {order.Address}", "MainViewModel");
                }
                
                // Aktualizuj NIP (vat)
                if (root.TryGetProperty("vat", out var vatProp) && vatProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Nip = vatProp.GetString();
                    Debug("Zaktualizowano NIP: {order.Nip}", "MainViewModel");
                }

                // Aktualizuj kraj - użyj kod ISO 2 aby znaleźć polską nazwę kraju
                JsonElement? iso2Prop = null;
                string? iso2Value = null;
                
                // Spróbuj najpierw payment_iso_code_2 (szczegóły zamówienia)
                if (root.TryGetProperty("payment_iso_code_2", out var paymentIso2Prop))
                {
                    iso2Prop = paymentIso2Prop;
                }
                // Fallback: iso_code_2 (lista zamówień)
                else if (root.TryGetProperty("iso_code_2", out var iso2PropFallback))
                {
                    iso2Prop = iso2PropFallback;
                }
                
                if (iso2Prop.HasValue && iso2Prop.Value.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    iso2Value = iso2Prop.Value.GetString()?.ToUpperInvariant();
                    order.IsoCode2 = iso2Value;
                    if (!string.IsNullOrWhiteSpace(iso2Value) && _countryMap.TryGetValue(iso2Value, out var polishName))
                    {
                        order.Country = polishName;
                        Debug("Zaktualizowano kraj (ISO {iso2Value}): {order.Country}", "MainViewModel");
                    }
                    else
                    {
                        // Jeśli nie znaleziono w mapie, użyj oryginalnej nazwy z API
                        if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                        {
                            order.Country = countryProp.GetString() ?? "";
                            Debug("Zaktualizowano kraj (oryginalna nazwa): {order.Country}", "MainViewModel");
                        }
                    }
                }
                else
                {
                    // Fallback: użyj oryginalnej nazwy z API
                    if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        order.Country = countryProp.GetString() ?? "";
                        Debug("Zaktualizowano kraj (oryginalna nazwa, brak ISO 2): {order.Country}", "MainViewModel");
                    }
                }
                
                // Aktualizuj kod ISO 3
                if (root.TryGetProperty("iso_code_3", out var iso3Prop) && iso3Prop.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.IsoCode3 = iso3Prop.GetString();
                    Debug("Zaktualizowano ISO Code 3: {order.IsoCode3}", "MainViewModel");
                }

                // Aktualizuj currency_value z API
                if (root.TryGetProperty("currency_value", out var currValueProp))
                {
                    double? currencyValue = null;
                    if (currValueProp.ValueKind == System.Text.Json.JsonValueKind.Number)
                    {
                        currencyValue = currValueProp.GetDouble();
                    }
                    else if (currValueProp.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        var currValueStr = currValueProp.GetString();
                        if (!string.IsNullOrWhiteSpace(currValueStr) && double.TryParse(currValueStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            currencyValue = parsedValue;
                        }
                    }
                    
                    if (currencyValue.HasValue)
                    {
                        order.CurrencyValue = currencyValue;
                        Debug("Zaktualizowano currency_value: {currencyValue.Value:F4}", "MainViewModel");
                    }
                }

                // Produkty
                try
                {
                    if (root.TryGetProperty("products", out var productsObj) && productsObj.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        if (productsObj.TryGetProperty("product", out var productArray))
                        {
                            var items = new List<Product>();
                            foreach (var p in productArray.EnumerateArray())
                            {
                                var product = new Product();
                                if (p.TryGetProperty("product_id", out var pid) && pid.ValueKind == System.Text.Json.JsonValueKind.String)
                                {
                                    product.ProductId = pid.GetString() ?? "";
                                }
                                if (p.TryGetProperty("name", out var pname)) product.Name = pname.GetString() ?? "";
                                if (p.TryGetProperty("model", out var pmodel)) product.Model = pmodel.GetString() ?? "";
                                if (p.TryGetProperty("quantity", out var pqty))
                                {
                                    var qtyStr = pqty.GetString();
                                    if (int.TryParse(qtyStr, out var q)) product.Quantity = q; else product.Quantity = 1;
                                }
                                if (p.TryGetProperty("price", out var pprice)) product.Price = pprice.GetString() ?? "";
                                if (p.TryGetProperty("total", out var ptotal)) product.Total = ptotal.GetString() ?? "";
                                if (p.TryGetProperty("tax", out var ptax)) product.Tax = ptax.GetString() ?? "";
                                if (p.TryGetProperty("taxrate", out var ptx)) product.TaxRate = ptx.GetString() ?? "";
                                if (p.TryGetProperty("discount", out var pdisc)) product.Discount = pdisc.GetString() ?? "";
                                
                                // Opcje produktu
                                if (p.TryGetProperty("options", out var optionsObj) && optionsObj.ValueKind == System.Text.Json.JsonValueKind.Object)
                                {
                                    if (optionsObj.TryGetProperty("option", out var optionArray))
                                    {
                                        var options = new List<ProductOption>();
                                        foreach (var opt in optionArray.EnumerateArray())
                                        {
                                            var option = new ProductOption();
                                            if (opt.TryGetProperty("name", out var optName)) option.Name = optName.GetString() ?? "";
                                            if (opt.TryGetProperty("value", out var optValue)) option.Value = optValue.GetString() ?? "";
                                            if (!string.IsNullOrWhiteSpace(option.Name) || !string.IsNullOrWhiteSpace(option.Value))
                                            {
                                                options.Add(option);
                                            }
                                        }
                                        if (options.Count > 0)
                                        {
                                            product.Options = options;
                                        }
                                    }
                                }
                                
                                items.Add(product);
                            }
                            order.Items = items;
                            Debug("Załadowano produkty: {items.Count}", "MainViewModel");
                        }
                    }
                }
                catch (Exception prodEx)
                {
                    Error(prodEx, "MainViewModel", "Błąd parsowania produktów");
                }

                // Totals - wyszukaj kupon i sub_total (API zwraca wartości netto dla kosztów)
                // Resetuj wartości przed parsowaniem, aby uniknąć kumulowania przy wielokrotnym wywołaniu
                order.GlsKgAmountNetto = null;
                order.GlsAmountNetto = null;
                
                // Flagi do sprawdzenia czy API zwraca wartości brutto dla konkretnych kosztów
                bool hasHandlingBrutto = false;
                bool hasGlsBrutto = false;
                bool hasGlsKgBrutto = false;
                bool hasShippingBrutto = false;
                bool hasCodFeeBrutto = false;
                
                try
                {
                    if (root.TryGetProperty("totals", out var totalsProp) && totalsProp.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        // Najpierw przejdź przez totals aby sprawdzić czy są pozycje brutto dla konkretnych kosztów
                        foreach (var totalEl in totalsProp.EnumerateArray())
                        {
                            if (totalEl.TryGetProperty("code", out var codeProp) && codeProp.ValueKind == System.Text.Json.JsonValueKind.String)
                            {
                                var code = codeProp.GetString();
                                Debug("Sprawdzam pozycję totals: code={code}", "MainViewModel");
                                // Sprawdź czy są pozycje brutto dla konkretnych kosztów (np. "handling_brutto", "shipping_brutto")
                                if (code == "handling_brutto" || code == "handling_bruto") hasHandlingBrutto = true;
                                if (code == "gls_brutto" || code == "gls_bruto") hasGlsBrutto = true;
                                if (code == "gls_kg_brutto" || code == "gls_kg_bruto") hasGlsKgBrutto = true;
                                if (code == "shipping_brutto" || code == "shipping_bruto") hasShippingBrutto = true;
                                if (code == "cod_fee_brutto" || code == "cod_fee_bruto") hasCodFeeBrutto = true;
                            }
                        }
                        
                        Info($"Wynik sprawdzenia brutto dla kosztów: handling={hasHandlingBrutto}, gls={hasGlsBrutto}, glsKg={hasGlsKgBrutto}, shipping={hasShippingBrutto}, codFee={hasCodFeeBrutto}", "MainViewModel");
                        
                        // Teraz parsuj wartości (wartości z API dla handling, gls, shipping, cod_fee są netto)
                        foreach (var totalEl in totalsProp.EnumerateArray())
                        {
                            if (totalEl.TryGetProperty("code", out var codeProp) && codeProp.ValueKind == System.Text.Json.JsonValueKind.String)
                            {
                                var code = codeProp.GetString();
                                double? valueParsed = null;
                                
                                // Spróbuj parsować jako string lub number
                                if (totalEl.TryGetProperty("value", out var valueProp))
                                {
                                    if (valueProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                    {
                                        var valueStr = valueProp.GetString();
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value))
                                        {
                                            valueParsed = value;
                                        }
                                    }
                                    else if (valueProp.ValueKind == System.Text.Json.JsonValueKind.Number)
                                    {
                                        valueParsed = valueProp.GetDouble();
                                    }
                                }

                                if (valueParsed.HasValue)
                                {
                                    if (code == "coupon")
                                    {
                                        // Wartość kuponu jest ujemna w API (np. -18.6665), więc użyj wartości bezwzględnej
                                        order.CouponAmount = Math.Abs(valueParsed.Value);
                                        // Parsuj tytuł kuponu (np. "Kupon (-15% , ASA5050)")
                                        if (totalEl.TryGetProperty("title", out var titleProp) && titleProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                        {
                                            order.CouponTitle = titleProp.GetString();
                                        }
                                        Debug("Znaleziono kupon: {order.CouponTitle} (wartość: {order.CouponAmount:F2})", "MainViewModel");
                                    }
                                    else if (code == "sub_total")
                                    {
                                        order.SubTotal = valueParsed.Value;
                                        Debug("Znaleziono sub_total: {order.SubTotal:F2}", "MainViewModel");
                                    }
                                    else if (code == "handling")
                                    {
                                        // Wartości z API są netto
                                        order.HandlingAmountNetto = valueParsed.Value;
                                        Debug("Znaleziono handling (netto): {order.HandlingAmountNetto:F2}", "MainViewModel");
                                    }
                                    else if (code == "gls")
                                    {
                                        // Sprawdź czy tytuł zawiera "kg" (case-insensitive)
                                        string? title = null;
                                        if (totalEl.TryGetProperty("title", out var titleProp) && titleProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                        {
                                            title = titleProp.GetString();
                                        }
                                        
                                        bool zawieraKg = !string.IsNullOrWhiteSpace(title) && title.Contains("kg", StringComparison.OrdinalIgnoreCase);
                                        
                                        if (zawieraKg)
                                        {
                                            // Sumuj pozycje gls z "kg" w tytule (wartości netto)
                                            order.GlsKgAmountNetto = (order.GlsKgAmountNetto ?? 0.0) + valueParsed.Value;
                                            Debug("Znaleziono gls z 'kg' (netto): {valueParsed.Value:F2} (tytuł: {title}), suma: {order.GlsKgAmountNetto:F2}", "MainViewModel");
                                        }
                                        else
                                        {
                                            // Zwykły gls - dodaj do GlsAmountNetto (lub sumuj jeśli wiele pozycji) - wartości netto
                                            order.GlsAmountNetto = (order.GlsAmountNetto ?? 0.0) + valueParsed.Value;
                                            Debug("Znaleziono gls (netto): {valueParsed.Value:F2} (tytuł: {title}), suma: {order.GlsAmountNetto:F2}", "MainViewModel");
                                        }
                                    }
                                    else if (code == "shipping")
                                    {
                                        // Wartości netto z API
                                        order.ShippingAmountNetto = valueParsed.Value;
                                        Debug("Znaleziono shipping (netto): {order.ShippingAmountNetto:F2}", "MainViewModel");
                                    }
                                    else if (code == "cod_fee")
                                    {
                                        // Wartości netto z API
                                        order.CodFeeAmountNetto = valueParsed.Value;
                                        Debug("Znaleziono cod_fee (netto): {order.CodFeeAmountNetto:F2}", "MainViewModel");
                                    }
                                    else if (code == "total")
                                    {
                                        order.Total = valueParsed.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                                        Debug("Znaleziono total: {order.Total}", "MainViewModel");
                                    }
                                    else if (code == "tax")
                                    {
                                        // Sprawdź czy tytuł zaczyna się od "VAT EU Export"
                                        string? title = null;
                                        if (totalEl.TryGetProperty("title", out var titleProp) && titleProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                        {
                                            title = titleProp.GetString();
                                        }
                                        
                                        if (!string.IsNullOrWhiteSpace(title) && title.StartsWith("VAT EU Export", StringComparison.OrdinalIgnoreCase))
                                        {
                                            order.UseEuVatRate = true;
                                            Debug("Znaleziono VAT EU Export w totals - ustawiono UseEuVatRate=true (tytuł: {title})", "MainViewModel");
                                        }
                                    }
                                    // Pomijamy pozycje "vat", "brutto" - są to pozycje dla całego zamówienia
                                    // Wartości kosztów (handling, gls, shipping, cod_fee) w totals są zawsze netto
                                }
                            }
                        }
                    }
                }
                catch (Exception totalsEx)
                {
                    Error(totalsEx, "MainViewModel", "Błąd parsowania totals");
                }
            }
            catch (Exception ex)
            {
                Error(ex, "MainViewModel", "Błąd parsowania szczegółów zamówienia");
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();
        
        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T?> _execute;
        private readonly Func<T?, bool>? _canExecute;

        public RelayCommand(Action<T?> execute, Func<T?, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter)
        {
            if (parameter is T t)
                return _canExecute?.Invoke(t) ?? true;
            return _canExecute?.Invoke(default) ?? true;
        }

        public void Execute(object? parameter)
        {
            if (parameter is T t)
                _execute(t);
            else
                _execute(default);
        }
    }
}

