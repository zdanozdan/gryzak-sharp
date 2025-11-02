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

        public ObservableCollection<Order> AllOrders { get; } = new ObservableCollection<Order>();
        public ObservableCollection<Order> FilteredOrders { get; } = new ObservableCollection<Order>();

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
                _searchText = value;
                OnPropertyChanged();
                FilterOrders();
            }
        }

        public string StatusFilter
        {
            get => _statusFilter;
            set
            {
                _statusFilter = value;
                OnPropertyChanged();
                FilterOrders();
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
            }
        }

        public string SubiektStatusText => IsSubiektActive ? "Subiekt GT aktywny (1 licencja)" : "";

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
        public ICommand ZwolnijLicencjeCommand { get; }
        public ICommand ClearSearchCommand { get; }

        public MainViewModel()
        {
            _configService = new ConfigService();
            _apiService = new ApiService(_configService);

            RefreshCommand = new RelayCommand(async () => await LoadOrdersAsync(true));
            ConfigureApiCommand = new RelayCommand(() => OpenConfigDialog());
            OpenSubiektSettingsCommand = new RelayCommand(() => OpenSubiektSettingsDialog());
            OrderSelectedCommand = new RelayCommand<Order>(order => OnOrderSelected(order));
            DodajZKCommand = new RelayCommand(() => DodajZK());
            ZwolnijLicencjeCommand = new RelayCommand(() => ZwolnijLicencje(), () => IsSubiektActive);
            ClearSearchCommand = new RelayCommand(() => { SearchText = ""; });

            // Zapisz się na event zmiany instancji Subiekta
            Services.SubiektService.InstancjaZmieniona += SubiektService_InstancjaZmieniona;
            
            // Sprawdź początkowy status
            var subiektService = new Services.SubiektService();
            IsSubiektActive = subiektService.CzyInstancjaAktywna();

            CheckApiConfiguration();
            
            // Upewnij się, że StatusFilter jest ustawione przed pierwszym ładowaniem
            StatusFilter = "Wszystkie statusy";
            
            LoadCountryMap();
            
            // Nie ładuj zamówień automatycznie - będzie to zrobione podczas splash screen
            // _ = LoadOrdersAsync(false);
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
            subiektService.OtworzOknoZK(nip, SelectedOrder?.Items, SelectedOrder?.CouponAmount, SelectedOrder?.SubTotal, SelectedOrder?.CouponTitle, SelectedOrder?.Id, SelectedOrder?.HandlingAmount, SelectedOrder?.ShippingAmount, SelectedOrder?.Currency, SelectedOrder?.CodFeeAmount, SelectedOrder?.Total, SelectedOrder?.GlsAmount, orderEmail, SelectedOrder?.Customer, SelectedOrder?.Phone, SelectedOrder?.Company, SelectedOrder?.Address, SelectedOrder?.PaymentAddress1, SelectedOrder?.PaymentAddress2, SelectedOrder?.PaymentPostcode, SelectedOrder?.PaymentCity, SelectedOrder?.Country, SelectedOrder?.IsoCode2);
            // Status zostanie zaktualizowany przez event InstancjaZmieniona
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
            });
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

                // Totals - wyszukaj kupon i sub_total (API zwraca już przeliczone wartości)
                try
                {
                    if (root.TryGetProperty("totals", out var totalsProp) && totalsProp.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
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
                                        order.HandlingAmount = valueParsed.Value;
                                        Debug("Znaleziono handling: {order.HandlingAmount:F2}", "MainViewModel");
                                    }
                                    else if (code == "gls")
                                    {
                                        order.GlsAmount = valueParsed.Value;
                                        Debug("Znaleziono gls: {order.GlsAmount:F2}", "MainViewModel");
                                    }
                                    else if (code == "shipping")
                                    {
                                        order.ShippingAmount = valueParsed.Value;
                                        Debug("Znaleziono shipping: {order.ShippingAmount:F2}", "MainViewModel");
                                    }
                                    else if (code == "cod_fee")
                                    {
                                        order.CodFeeAmount = valueParsed.Value;
                                        Debug("Znaleziono cod_fee: {order.CodFeeAmount:F2}", "MainViewModel");
                                    }
                                    else if (code == "total")
                                    {
                                        order.Total = valueParsed.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                                        Debug("Znaleziono total: {order.Total}", "MainViewModel");
                                    }
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

