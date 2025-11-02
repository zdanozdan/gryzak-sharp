using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using Gryzak.Models;
using Gryzak.ViewModels;
using Gryzak.Converters;
using static Gryzak.Services.Logger;

namespace Gryzak.Views
{
    public partial class OrderDetailsDialog : Window, INotifyPropertyChanged
    {
        private string _orderId = "";
        private string _customerName = "";
        private DateTime _orderDate = DateTime.Now;
        private string? _customerCompany;
        private string? _customerNip;
        private string? _customerAddress;
        private string _customerCountry = "";
        private List<Product> _products = new List<Product>();
        private ObservableCollection<OrderTotal> _orderTotals = new ObservableCollection<OrderTotal>();
        private string _total = "0.00";
        private string _currency = "PLN";
        private Order? _order;
        private MainViewModel? _mainViewModel;

        public OrderDetailsDialog(Order order, MainViewModel? mainViewModel = null)
        {
            InitializeComponent();
            
            _order = order;
            _mainViewModel = mainViewModel;
            
            // Inicjalizuj OrderTotals
            OrderTotals = new ObservableCollection<OrderTotal>();
            
            // Wczytaj dane z zamówienia
            OrderId = order.Id;
            CustomerName = order.Customer;
            OrderDate = order.Date;
            CustomerCompany = order.Company;
            CustomerNip = order.Nip;
            CustomerAddress = order.Address;
            CustomerCountry = order.CountryWithIso3;
            Products = order.Items ?? new List<Product>();
            Total = order.Total;
            Currency = order.Currency;
            
            DataContext = this;
            
            // Zawsze pobierz szczegóły z API aby zaktualizować totals
            LoadOrderDetailsAsync();
        }

        public string OrderId
        {
            get => _orderId;
            set { _orderId = value; OnPropertyChanged(); }
        }

        public string CustomerName
        {
            get => _customerName;
            set { _customerName = value; OnPropertyChanged(); }
        }

        public DateTime OrderDate
        {
            get => _orderDate;
            set { _orderDate = value; OnPropertyChanged(); }
        }

        public string? CustomerCompany
        {
            get => _customerCompany;
            set { _customerCompany = value; OnPropertyChanged(); }
        }

        public string? CustomerNip
        {
            get => _customerNip;
            set { _customerNip = value; OnPropertyChanged(); }
        }

        public string? CustomerAddress
        {
            get => _customerAddress;
            set { _customerAddress = value; OnPropertyChanged(); }
        }

        public string CustomerCountry
        {
            get => _customerCountry;
            set { _customerCountry = value; OnPropertyChanged(); }
        }

        public List<Product> Products
        {
            get => _products;
            set 
            { 
                _products = value; 
                OnPropertyChanged(); 
                OnPropertyChanged(nameof(HasProducts));
                OnPropertyChanged(nameof(TotalQuantity));
                OnPropertyChanged(nameof(TotalTax));
                OnPropertyChanged(nameof(TotalPriceNet));
                OnPropertyChanged(nameof(TotalPriceGross));
                OnPropertyChanged(nameof(TotalSum));
            }
        }

        public bool HasProducts => Products != null && Products.Count > 0;

        public int TotalQuantity
        {
            get
            {
                if (Products == null) return 0;
                return Products.Sum(p => p.Quantity);
            }
        }

        public string TotalTax
        {
            get
            {
                if (Products == null || Products.Count == 0) return "0.00";
                double sum = 0.0;
                foreach (var p in Products)
                {
                    if (!string.IsNullOrWhiteSpace(p.Tax))
                    {
                        if (double.TryParse(p.Tax, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var tax))
                        {
                            sum += tax * p.Quantity;
                        }
                    }
                }
                return sum.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public string TotalPriceNet
        {
            get
            {
                if (Products == null || Products.Count == 0) return "0.00";
                double sum = 0.0;
                foreach (var p in Products)
                {
                    if (!string.IsNullOrWhiteSpace(p.Price))
                    {
                        if (double.TryParse(p.Price, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var price))
                        {
                            sum += price * p.Quantity;
                        }
                    }
                }
                return sum.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public string TotalPriceGross
        {
            get
            {
                if (Products == null || Products.Count == 0) return "0.00";
                double sum = 0.0;
                foreach (var p in Products)
                {
                    double price = 0.0;
                    double tax = 0.0;
                    
                    if (!string.IsNullOrWhiteSpace(p.Price))
                    {
                        double.TryParse(p.Price, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out price);
                    }
                    
                    if (!string.IsNullOrWhiteSpace(p.Tax))
                    {
                        double.TryParse(p.Tax, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out tax);
                    }
                    
                    sum += (price + tax) * p.Quantity;
                }
                return sum.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public string TotalSum
        {
            get
            {
                if (Products == null || Products.Count == 0) return "0.00";
                double sum = 0.0;
                foreach (var p in Products)
                {
                    if (!string.IsNullOrWhiteSpace(p.Total))
                    {
                        if (double.TryParse(p.Total, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var total))
                        {
                            sum += total;
                        }
                    }
                }
                return sum.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public string TotalSumGross
        {
            get
            {
                if (Products == null || Products.Count == 0) return "0.00";
                double sum = 0.0;
                foreach (var p in Products)
                {
                    double price = 0.0;
                    double tax = 0.0;
                    
                    if (!string.IsNullOrWhiteSpace(p.Price))
                    {
                        double.TryParse(p.Price, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out price);
                    }
                    
                    if (!string.IsNullOrWhiteSpace(p.Tax))
                    {
                        double.TryParse(p.Tax, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out tax);
                    }
                    
                    sum += (price + tax) * p.Quantity;
                }
                return sum.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public string Total
        {
            get => _total;
            set { _total = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalAmount)); }
        }

        public string Currency
        {
            get => _currency;
            set { _currency = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalAmount)); }
        }

        public string TotalAmount => $"{Total} {Currency}";

        public ObservableCollection<OrderTotal> OrderTotals
        {
            get => _orderTotals;
            set 
            { 
                _orderTotals = value; 
                OnPropertyChanged(); 
            }
        }

        private async void LoadOrderDetailsAsync()
        {
            if (_order == null) return;

            try
            {
                var apiService = new Services.ApiService(new Services.ConfigService());
                var detailsJson = await apiService.GetOrderDetailsAsync(_order.Id);
                
                // Zaktualizuj zamówienie z szczegółami używając tej samej logiki co MainViewModel
                UpdateOrderFromDetails(detailsJson);
            }
            catch (Exception ex)
            {
                Error(ex, "OrderDetailsDialog", "Błąd ładowania szczegółów");
            }
        }

        private void UpdateOrderFromDetails(string detailsJson)
        {
            if (_order == null) return;
            
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(detailsJson);
                var root = doc.RootElement;

                // Aktualizuj nazwę firmy (payment_company)
                if (root.TryGetProperty("payment_company", out var companyProp) && companyProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    var companyValue = companyProp.GetString();
                    if (!string.IsNullOrWhiteSpace(companyValue))
                    {
                        // Decoduj HTML entities dwukrotnie (bo API zwraca podwójnie zakodowane encje)
                        companyValue = System.Net.WebUtility.HtmlDecode(companyValue);
                        companyValue = System.Net.WebUtility.HtmlDecode(companyValue);
                        _order.Company = companyValue;
                        CustomerCompany = companyValue;
                        Debug("Zaktualizowano firmę: {CustomerCompany}", "OrderDetailsDialog");
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
                    
                    var fullAddress = string.Join(", ", addressParts.Where(s => !string.IsNullOrWhiteSpace(s)));
                    _order.Address = string.IsNullOrWhiteSpace(fullAddress) ? null : fullAddress;
                    CustomerAddress = _order.Address;
                    Debug("Zaktualizowano adres: {CustomerAddress}", "OrderDetailsDialog");
                }

                // Aktualizuj NIP (vat)
                if (root.TryGetProperty("vat", out var vatProp) && vatProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    _order.Nip = vatProp.GetString();
                    CustomerNip = _order.Nip;
                    Debug("Zaktualizowano NIP: {CustomerNip}", "OrderDetailsDialog");
                }

                // Aktualizuj kraj - użyj kod ISO 2 aby znaleźć polską nazwę kraju
                System.Text.Json.JsonElement? iso2Prop = null;
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
                    _order.IsoCode2 = iso2Value;
                    
                    // Jeśli mamy MainViewModel, użyj jego mapy krajów
                    if (_mainViewModel != null)
                    {
                        var countryMap = _mainViewModel.GetCountryMap();
                        if (!string.IsNullOrWhiteSpace(iso2Value) && countryMap.TryGetValue(iso2Value, out var polishName))
                        {
                            _order.Country = polishName;
                            CustomerCountry = _order.CountryWithIso3;
                            Debug("Zaktualizowano kraj (ISO {iso2Value}): {_order.Country}", "OrderDetailsDialog");
                        }
                        else
                        {
                            // Jeśli nie znaleziono w mapie, użyj oryginalnej nazwy z API
                            if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                            {
                                _order.Country = countryProp.GetString() ?? "";
                                CustomerCountry = _order.CountryWithIso3;
                                Debug("Zaktualizowano kraj (oryginalna nazwa): {_order.Country}", "OrderDetailsDialog");
                            }
                        }
                    }
                    else
                    {
                        // Jeśli nie mamy MainViewModel, użyj oryginalnej nazwy
                        if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                        {
                            _order.Country = countryProp.GetString() ?? "";
                            CustomerCountry = _order.CountryWithIso3;
                            Debug("Zaktualizowano kraj (oryginalna nazwa, brak MainViewModel): {_order.Country}", "OrderDetailsDialog");
                        }
                    }
                }
                else
                {
                    // Fallback: użyj oryginalnej nazwy z API
                    if (root.TryGetProperty("payment_country", out var countryProp) && countryProp.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        _order.Country = countryProp.GetString() ?? "";
                        CustomerCountry = _order.CountryWithIso3;
                        Debug("Zaktualizowano kraj (oryginalna nazwa, brak ISO 2): {_order.Country}", "OrderDetailsDialog");
                    }
                }
                
                // Aktualizuj kod ISO 3
                if (root.TryGetProperty("iso_code_3", out var iso3Prop) && iso3Prop.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    _order.IsoCode3 = iso3Prop.GetString();
                    CustomerCountry = _order.CountryWithIso3;
                    Debug("Zaktualizowano ISO Code 3: {_order.IsoCode3}", "OrderDetailsDialog");
                }

                // Aktualizuj walutę
                if (root.TryGetProperty("currency_code", out var currencyProp) && currencyProp.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    var currencyValue = currencyProp.GetString() ?? "PLN";
                    Currency = currencyValue;
                    _order.Currency = currencyValue;
                    Debug("Zaktualizowano walutę: {Currency}", "OrderDetailsDialog");
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
                            Products = items;
                            _order.Items = items;
                            // Powiadom o zmianach w podsumowaniu
                            OnPropertyChanged(nameof(TotalQuantity));
                            OnPropertyChanged(nameof(TotalTax));
                            OnPropertyChanged(nameof(TotalPriceNet));
                            OnPropertyChanged(nameof(TotalPriceGross));
                            OnPropertyChanged(nameof(TotalSum));
                            OnPropertyChanged(nameof(TotalSumGross));
                            Debug("Załadowano produkty: {items.Count}", "OrderDetailsDialog");
                        }
                    }
                }
                catch (Exception prodEx)
                {
                    Error(prodEx, "OrderDetailsDialog", "Błąd parsowania produktów");
                }

                // Totals - parsuj wszystkie totals z API
                try
                {
                    var totalsList = new List<OrderTotal>();
                    
                    if (root.TryGetProperty("totals", out var totalsProp) && totalsProp.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        foreach (var totalEl in totalsProp.EnumerateArray())
                        {
                            var orderTotal = new OrderTotal();
                            
                            // Parsuj code
                            if (totalEl.TryGetProperty("code", out var codeProp) && codeProp.ValueKind == System.Text.Json.JsonValueKind.String)
                            {
                                orderTotal.Code = codeProp.GetString() ?? "";
                            }
                            
                            // Parsuj title
                            if (totalEl.TryGetProperty("title", out var titleProp) && titleProp.ValueKind == System.Text.Json.JsonValueKind.String)
                            {
                                orderTotal.Title = titleProp.GetString() ?? "";
                            }
                            
                            // Parsuj value
                            if (totalEl.TryGetProperty("value", out var valueProp))
                            {
                                if (valueProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                {
                                    var valueStr = valueProp.GetString();
                                    if (double.TryParse(valueStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value))
                                    {
                                        orderTotal.Value = value;
                                    }
                                }
                                else if (valueProp.ValueKind == System.Text.Json.JsonValueKind.Number)
                                {
                                    orderTotal.Value = valueProp.GetDouble();
                                }
                            }
                            
                            // Parsuj sort_order
                            if (totalEl.TryGetProperty("sort_order", out var sortProp))
                            {
                                if (sortProp.ValueKind == System.Text.Json.JsonValueKind.String)
                                {
                                    var sortStr = sortProp.GetString();
                                    if (int.TryParse(sortStr, out var sortOrder))
                                    {
                                        orderTotal.SortOrder = sortOrder;
                                    }
                                }
                                else if (sortProp.ValueKind == System.Text.Json.JsonValueKind.Number)
                                {
                                    orderTotal.SortOrder = sortProp.GetInt32();
                                }
                            }
                            
                            totalsList.Add(orderTotal);
                            
                            // Loguj znalezione total
                            Debug("Znaleziono total: code={orderTotal.Code}, title={orderTotal.Title}, value={orderTotal.Value:F2}, sort={orderTotal.SortOrder}", "OrderDetailsDialog");
                            
                            // Ustaw Total jeśli to główny total
                            if (orderTotal.Code == "total")
                            {
                                Total = orderTotal.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                                _order.Total = Total;
                            }
                        }
                        
                        // Sortuj po sort_order
                        totalsList.Sort((a, b) => a.SortOrder.CompareTo(b.SortOrder));
                        
                        // Ustaw kolekcję
                        OrderTotals = new ObservableCollection<OrderTotal>(totalsList);
                        Debug("Załadowano {totalsList.Count} totals", "OrderDetailsDialog");
                    }
                }
                catch (Exception totalsEx)
                {
                    Error(totalsEx, "OrderDetailsDialog", "Błąd parsowania totals");
                }
            }
            catch (Exception ex)
            {
                Error(ex, "OrderDetailsDialog", "Błąd parsowania szczegółów zamówienia");
            }
        }

        private void OpenZKButton_Click(object sender, RoutedEventArgs e)
        {
            Debug("Przycisk 'Otwórz ZK' został kliknięty", "OrderDetailsDialog");
            
            // Zapisz dane lokalnie przed zamknięciem okna
            var order = _order;
            var mainViewModel = _mainViewModel;
            
            // Zamknij okno najpierw, aby nie przeszkadzało w kolejnych dialogach
            Debug("Zamykanie okna szczegółów zamówienia", "OrderDetailsDialog");
            DialogResult = true;
            Close();
            
            // Wykonaj akcję po zamknięciu okna (używamy Dispatcher aplikacji bo okno już zamknięte)
            if (order != null && mainViewModel != null)
            {
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    Debug("Uruchamianie komendy DodajZK po zamknięciu okna", "OrderDetailsDialog");
                    // Zaznacz zamówienie i otwórz ZK
                    mainViewModel.SelectedOrder = order;
                    mainViewModel.DodajZKCommand.Execute(null);
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

