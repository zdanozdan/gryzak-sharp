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
using Gryzak.Models;
using Gryzak.Services;

namespace Gryzak.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ApiService _apiService;
        private readonly ConfigService _configService;
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
            
            _ = LoadOrdersAsync(false);
        }

        private void CheckApiConfiguration()
        {
            var config = _configService.LoadConfig();
            IsApiConfigured = !string.IsNullOrWhiteSpace(config.ApiUrl);
        }

        private async Task LoadOrdersAsync(bool reset = false)
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
                    UpdateProgress("Ładowanie zamówień...", 10);
                }
                else
                {
                    IsLoadingMore = true;
                }

                var orders = await _apiService.LoadOrdersAsync(_currentPage);
                
                Console.WriteLine($"[Gryzak] Załadowano stronę {_currentPage}: {orders.Count} zamówień");
                
                if (isFirstPage)
                {
                    UpdateProgress("Przetwarzanie danych...", 50);
                }

                // Jeśli API zwraca mniej niż 20 zamówień lub 0, uznaj to za ostatnią stronę
                if (orders.Count < 20)
                {
                    _hasMorePages = false;
                    OnPropertyChanged(nameof(HasMorePages));
                    Console.WriteLine($"[Gryzak] Ostatnia strona - zwrócono {orders.Count} zamówień");
                }

                if (orders.Count == 0)
                {
                    Console.WriteLine($"[Gryzak] Brak zamówień na stronie {_currentPage}");
                }
                else
                {
                    foreach (var order in orders)
                    {
                        // Sprawdź czy zamówienie nie jest już w liście (duplikaty)
                        if (!AllOrders.Any(o => o.Id == order.Id))
                        {
                            AllOrders.Add(order);
                            Console.WriteLine($"[Gryzak] Dodano zamówienie: {order.Id} - {order.Customer}");
                            
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
                    UpdateProgress("Ładowanie zakończone", 100);
                    await Task.Delay(500); // Pokaż 100% przez 500ms
                    HideProgress();
                }
                
                Console.WriteLine($"[Gryzak] Przetworzonych zamówień: {AllOrders.Count}, przefiltrowanych: {FilteredOrders.Count}, HasOrders: {HasOrders}, HasMorePages: {_hasMorePages}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Gryzak] Błąd ładowania: {ex.Message}");
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
                Console.WriteLine($"[Gryzak] IsLoading: {IsLoading}, IsLoadingMore: {IsLoadingMore}");
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
            
            Console.WriteLine($"[Gryzak] FilterOrders - StatusFilter: '{StatusFilter}', SearchText: '{SearchText}', AllOrders.Count: {AllOrders.Count}");
            
            if (AllOrders.Count > 0)
            {
                Console.WriteLine($"[Gryzak] Przykładowy status pierwszego zamówienia: '{AllOrders[0].Status}'");
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
            Console.WriteLine($"[Gryzak] Po filtrowaniu: {ordersList.Count} zamówień");

            foreach (var order in ordersList)
            {
                FilteredOrders.Add(order);
            }

            Console.WriteLine($"[Gryzak] FilteredOrders.Count: {FilteredOrders.Count}");
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
            
            // Jeśli NIP nie jest dostępny, pokaż ostrzeżenie
            if (string.IsNullOrWhiteSpace(nip))
            {
                MessageBox.Show(
                    "Wybrane zamówienie nie posiada numeru NIP.\n\n" +
                    "Dokument ZK zostanie otwarty bez przypisanego kontrahenta.",
                    "Brak NIP",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            
            var subiektService = new Services.SubiektService();
            subiektService.OtworzOknoZK(nip, SelectedOrder?.Items, SelectedOrder?.CouponAmount, SelectedOrder?.SubTotal, SelectedOrder?.CouponTitle, SelectedOrder?.Id, SelectedOrder?.HandlingAmount, SelectedOrder?.ShippingAmount, SelectedOrder?.Currency, SelectedOrder?.CodFeeAmount, SelectedOrder?.Total, SelectedOrder?.GlsAmount);
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
                Console.WriteLine("\n=== ZAZNACZONO ZAMÓWIENIE ===");
                Console.WriteLine($"ID Zamówienia: {order.Id}");
                Console.WriteLine($"Klient: {order.Customer}");
                Console.WriteLine($"Email: {order.Email}");
                Console.WriteLine($"Telefon: {order.Phone}");
                Console.WriteLine($"Firma: {order.Company ?? "Brak"}");
                Console.WriteLine($"NIP: {order.Nip ?? "Brak"}");
                Console.WriteLine($"Adres: {order.Address ?? "Brak"}");
                Console.WriteLine($"Status: {order.Status}");
                Console.WriteLine($"Status Płatności: {order.PaymentStatus}");
                Console.WriteLine($"Wartość: {order.Total} {order.Currency}");
                Console.WriteLine($"Data: {order.Date:dd.MM.yyyy HH:mm}");
                Console.WriteLine($"Obsługuje: {order.AssignedTo}");
                Console.WriteLine($"Liczba produktów: {order.Items?.Count ?? 0}");
                Console.WriteLine("==========================\n");
                
                // Pobierz szczegóły zamówienia z API
                LoadOrderDetailsAsync(order.Id);
            }
        }
        
        private async void LoadOrderDetailsAsync(string orderId)
        {
            try
            {
                Console.WriteLine($"[MainViewModel] Ładowanie szczegółów zamówienia {orderId}...");
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
                
                Console.WriteLine("\n=== SZCZEGÓŁY ZAMÓWIENIA Z API ===");
                Console.WriteLine(formattedJson);
                Console.WriteLine("===================================\n");
                
                // Zaktualizuj zamówienie w liście
                UpdateOrderFromDetails(SelectedOrder, detailsJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MainViewModel] Błąd ładowania szczegółów zamówienia: {ex.Message}");
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
                    Console.WriteLine($"[MainViewModel] Zaktualizowano email: {order.Email}");
                }
                
                // Aktualizuj telefon
                if (root.TryGetProperty("telephone", out var telProp) && telProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Phone = telProp.GetString() ?? "Brak telefonu";
                    Console.WriteLine($"[MainViewModel] Zaktualizowano telefon: {order.Phone}");
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
                        Console.WriteLine($"[MainViewModel] Zaktualizowano firmę: {order.Company}");
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
                    Console.WriteLine($"[MainViewModel] Zaktualizowano adres: {order.Address}");
                }
                
                // Aktualizuj NIP (vat)
                if (root.TryGetProperty("vat", out var vatProp) && vatProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Nip = vatProp.GetString();
                    Console.WriteLine($"[MainViewModel] Zaktualizowano NIP: {order.Nip}");
                }

                // Aktualizuj kraj
                if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.Country = countryProp.GetString() ?? "";
                    Console.WriteLine($"[MainViewModel] Zaktualizowano kraj: {order.Country}");
                }
                
                // Aktualizuj kod ISO 3
                if (root.TryGetProperty("iso_code_3", out var iso3Prop) && iso3Prop.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    order.IsoCode3 = iso3Prop.GetString();
                    Console.WriteLine($"[MainViewModel] Zaktualizowano ISO Code 3: {order.IsoCode3}");
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
                            Console.WriteLine($"[MainViewModel] Załadowano produkty: {items.Count}");
                        }
                    }
                }
                catch (Exception prodEx)
                {
                    Console.WriteLine($"[MainViewModel] Błąd parsowania produktów: {prodEx.Message}");
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
                                        Console.WriteLine($"[MainViewModel] Znaleziono kupon: {order.CouponTitle} (wartość: {order.CouponAmount:F2})");
                                    }
                                    else if (code == "sub_total")
                                    {
                                        order.SubTotal = valueParsed.Value;
                                        Console.WriteLine($"[MainViewModel] Znaleziono sub_total: {order.SubTotal:F2}");
                                    }
                                    else if (code == "handling")
                                    {
                                        order.HandlingAmount = valueParsed.Value;
                                        Console.WriteLine($"[MainViewModel] Znaleziono handling: {order.HandlingAmount:F2}");
                                    }
                                    else if (code == "gls")
                                    {
                                        order.GlsAmount = valueParsed.Value;
                                        Console.WriteLine($"[MainViewModel] Znaleziono gls: {order.GlsAmount:F2}");
                                    }
                                    else if (code == "shipping")
                                    {
                                        order.ShippingAmount = valueParsed.Value;
                                        Console.WriteLine($"[MainViewModel] Znaleziono shipping: {order.ShippingAmount:F2}");
                                    }
                                    else if (code == "cod_fee")
                                    {
                                        order.CodFeeAmount = valueParsed.Value;
                                        Console.WriteLine($"[MainViewModel] Znaleziono cod_fee: {order.CodFeeAmount:F2}");
                                    }
                                    else if (code == "total")
                                    {
                                        order.Total = valueParsed.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                                        Console.WriteLine($"[MainViewModel] Znaleziono total: {order.Total}");
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception totalsEx)
                {
                    Console.WriteLine($"[MainViewModel] Błąd parsowania totals: {totalsEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MainViewModel] Błąd parsowania szczegółów zamówienia: {ex.Message}");
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

