using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace Gryzak.Models
{
    public class Order : INotifyPropertyChanged
    {
        private bool _isSelected;

        private string _id = "";
        private string _customer = "";
        private string _email = "Brak email";
        private string _phone = "Brak telefonu";
        private string? _company;
        private string? _nip;
        private string? _address;
        private string? _paymentAddress1;
        private string? _paymentAddress2;
        private string? _paymentPostcode;
        private string? _paymentCity;
        private string? _shippingFirstname;
        private string? _shippingLastname;
        private string? _shippingCompany;
        private string? _shippingAddress1;
        private string? _shippingAddress2;
        private string? _shippingPostcode;
        private string? _shippingCity;
        
        public string Id 
        { 
            get => _id; 
            set 
            { 
                _id = value; 
                OnPropertyChanged();
                OnPropertyChanged(nameof(DisplayId)); // Powiadom o zmianie DisplayId
            } 
        }
        public string Customer { get => _customer; set { _customer = value; OnPropertyChanged(); } }
        public string Email { get => _email; set { _email = value; OnPropertyChanged(); } }
        public string Phone { get => _phone; set { _phone = value; OnPropertyChanged(); } }
        public string? Company { get => _company; set { _company = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsShippingDifferentFromPayment)); } }
        public string? Nip { get => _nip; set { _nip = value; OnPropertyChanged(); } }
        public string? Address { get => _address; set { _address = value; OnPropertyChanged(); } }
        public string? PaymentAddress1 { get => _paymentAddress1; set { _paymentAddress1 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsShippingDifferentFromPayment)); } }
        public string? PaymentAddress2 { get => _paymentAddress2; set { _paymentAddress2 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsShippingDifferentFromPayment)); } }
        public string? PaymentPostcode { get => _paymentPostcode; set { _paymentPostcode = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsShippingDifferentFromPayment)); } }
        public string? PaymentCity { get => _paymentCity; set { _paymentCity = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsShippingDifferentFromPayment)); OnPropertyChanged(nameof(IsShippingCityDifferent)); OnPropertyChanged(nameof(ShippingDifferenceSummary)); } }

        public string? ShippingFirstname
        {
            get => _shippingFirstname;
            set { _shippingFirstname = value; OnPropertyChanged(); OnPropertyChanged(nameof(ShippingDisplayName)); OnPropertyChanged(nameof(HasShippingAddress)); }
        }
        public string? ShippingLastname
        {
            get => _shippingLastname;
            set { _shippingLastname = value; OnPropertyChanged(); OnPropertyChanged(nameof(ShippingDisplayName)); OnPropertyChanged(nameof(HasShippingAddress)); }
        }
        public string? ShippingCompany
        {
            get => _shippingCompany;
            set
            {
                _shippingCompany = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShippingDisplayName));
                OnPropertyChanged(nameof(HasShippingAddress));
                OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
            }
        }
        public string? ShippingAddress1
        {
            get => _shippingAddress1;
            set
            {
                _shippingAddress1 = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShippingDisplayStreet));
                OnPropertyChanged(nameof(ShippingDisplayAddress));
                OnPropertyChanged(nameof(HasShippingAddress));
                OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
            }
        }
        public string? ShippingAddress2
        {
            get => _shippingAddress2;
            set
            {
                _shippingAddress2 = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShippingDisplayStreet));
                OnPropertyChanged(nameof(ShippingDisplayAddress));
                OnPropertyChanged(nameof(HasShippingAddress));
                OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
            }
        }
        public string? ShippingPostcode
        {
            get => _shippingPostcode;
            set
            {
                _shippingPostcode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShippingDisplayCityLine));
                OnPropertyChanged(nameof(ShippingDisplayAddress));
                OnPropertyChanged(nameof(HasShippingAddress));
                OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
            }
        }
        public string? ShippingCity
        {
            get => _shippingCity;
            set
            {
                _shippingCity = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShippingDisplayCityLine));
                OnPropertyChanged(nameof(ShippingDisplayAddress));
                OnPropertyChanged(nameof(HasShippingAddress));
                OnPropertyChanged(nameof(IsShippingCityDifferent));
                OnPropertyChanged(nameof(ShippingDifferenceSummary));
                OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
            }
        }

        /// <summary>Czy API zwróciło jakikolwiek adres wysyłki.</summary>
        public bool HasShippingAddress =>
            !string.IsNullOrWhiteSpace(ShippingAddress1)
            || !string.IsNullOrWhiteSpace(ShippingAddress2)
            || !string.IsNullOrWhiteSpace(ShippingPostcode)
            || !string.IsNullOrWhiteSpace(ShippingCity)
            || !string.IsNullOrWhiteSpace(ShippingCompany)
            || !string.IsNullOrWhiteSpace(ShippingFirstname)
            || !string.IsNullOrWhiteSpace(ShippingLastname);

        /// <summary>Nazwa odbiorcy / firmy do wyświetlenia w panelu adresu wysyłki.</summary>
        public string ShippingDisplayName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(ShippingCompany))
                    return ShippingCompany.Trim();
                return ShippingPersonName;
            }
        }

        /// <summary>Imię i nazwisko odbiorcy wysyłki.</summary>
        public string ShippingPersonName => $"{ShippingFirstname} {ShippingLastname}".Trim();

        /// <summary>Firma wysyłki (do osobnego wyświetlenia).</summary>
        public string ShippingCompanyDisplay => ShippingCompany?.Trim() ?? "";

        /// <summary>Ulica adresu wysyłki (bez kodu/miasta).</summary>
        public string ShippingDisplayStreet
        {
            get
            {
                var parts = new List<string>();
                if (!string.IsNullOrWhiteSpace(ShippingAddress1))
                    parts.Add(ShippingAddress1.Trim());
                if (!string.IsNullOrWhiteSpace(ShippingAddress2))
                    parts.Add(ShippingAddress2.Trim());
                return string.Join(", ", parts);
            }
        }

        /// <summary>Kod i miasto adresu wysyłki.</summary>
        public string ShippingDisplayCityLine => $"{ShippingPostcode} {ShippingCity}".Trim();

        /// <summary>Ulica + kod i miasto adresu wysyłki.</summary>
        public string ShippingDisplayAddress
        {
            get
            {
                var parts = new List<string>();
                if (!string.IsNullOrWhiteSpace(ShippingDisplayStreet))
                    parts.Add(ShippingDisplayStreet);
                if (!string.IsNullOrWhiteSpace(ShippingDisplayCityLine))
                    parts.Add(ShippingDisplayCityLine);
                return string.Join(", ", parts);
            }
        }

        public bool IsShippingCompanyDifferent =>
            HasShippingAddress
            && !string.Equals(NormalizeAddressPart(Company), NormalizeAddressPart(ShippingCompany), StringComparison.Ordinal);

        public bool IsShippingStreetDifferent =>
            HasShippingAddress
            && !string.Equals(
                NormalizeAddressPart($"{PaymentAddress1} {PaymentAddress2}"),
                NormalizeAddressPart($"{ShippingAddress1} {ShippingAddress2}"),
                StringComparison.Ordinal);

        public bool IsShippingPostcodeDifferent =>
            HasShippingAddress
            && !string.Equals(NormalizeAddressPart(PaymentPostcode), NormalizeAddressPart(ShippingPostcode), StringComparison.Ordinal);

        public bool IsShippingCityDifferent =>
            HasShippingAddress
            && !string.Equals(NormalizeAddressPart(PaymentCity), NormalizeAddressPart(ShippingCity), StringComparison.Ordinal);

        /// <summary>Lista różniących się pól, np. „firma, ulica”.</summary>
        public string ShippingDifferenceSummary
        {
            get
            {
                var parts = new List<string>();
                if (IsShippingCompanyDifferent) parts.Add("firma");
                if (IsShippingStreetDifferent) parts.Add("ulica");
                if (IsShippingPostcodeDifferent) parts.Add("kod pocztowy");
                if (IsShippingCityDifferent) parts.Add("miasto");
                return string.Join(", ", parts);
            }
        }

        /// <summary>
        /// True gdy adres wysyłki różni się od płatności (firma, ulica, kod, miasto).
        /// </summary>
        public bool IsShippingDifferentFromPayment =>
            IsShippingCompanyDifferent
            || IsShippingStreetDifferent
            || IsShippingPostcodeDifferent
            || IsShippingCityDifferent;

        /// <summary>
        /// Odczytuje pola payment_* i shipping_* ze szczegółów zamówienia OpenCart.
        /// </summary>
        public void ApplyAddressesFromApi(JsonElement root)
        {
            // Odbiorca płatności (imię/nazwisko z payment_* jeśli dostępne)
            var paymentFirst = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_firstname")));
            var paymentLast = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_lastname")));
            var paymentName = $"{paymentFirst} {paymentLast}".Trim();
            if (string.IsNullOrWhiteSpace(paymentName))
            {
                paymentFirst = NullIfEmpty(DecodeHtml(GetStringProp(root, "firstname")));
                paymentLast = NullIfEmpty(DecodeHtml(GetStringProp(root, "lastname")));
                paymentName = $"{paymentFirst} {paymentLast}".Trim();
            }
            if (!string.IsNullOrWhiteSpace(paymentName))
                Customer = paymentName;

            Company = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_company")));
            PaymentAddress1 = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_address_1")));
            PaymentAddress2 = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_address_2")));
            PaymentPostcode = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_postcode")));
            PaymentCity = NullIfEmpty(DecodeHtml(GetStringProp(root, "payment_city")));

            var paymentParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(PaymentAddress1)) paymentParts.Add(PaymentAddress1);
            if (!string.IsNullOrWhiteSpace(PaymentAddress2)) paymentParts.Add(PaymentAddress2);
            var paymentCityLine = $"{PaymentPostcode} {PaymentCity}".Trim();
            if (!string.IsNullOrWhiteSpace(paymentCityLine)) paymentParts.Add(paymentCityLine);
            Address = paymentParts.Count > 0 ? string.Join(", ", paymentParts) : null;

            ShippingFirstname = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_firstname")));
            ShippingLastname = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_lastname")));
            ShippingCompany = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_company")));
            ShippingAddress1 = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_address_1")));
            ShippingAddress2 = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_address_2")));
            ShippingPostcode = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_postcode")));
            ShippingCity = NullIfEmpty(DecodeHtml(GetStringProp(root, "shipping_city")));

            OnPropertyChanged(nameof(HasShippingAddress));
            OnPropertyChanged(nameof(ShippingDisplayName));
            OnPropertyChanged(nameof(ShippingPersonName));
            OnPropertyChanged(nameof(ShippingCompanyDisplay));
            OnPropertyChanged(nameof(ShippingDisplayStreet));
            OnPropertyChanged(nameof(ShippingDisplayCityLine));
            OnPropertyChanged(nameof(ShippingDisplayAddress));
            OnPropertyChanged(nameof(IsShippingCompanyDifferent));
            OnPropertyChanged(nameof(IsShippingStreetDifferent));
            OnPropertyChanged(nameof(IsShippingPostcodeDifferent));
            OnPropertyChanged(nameof(IsShippingCityDifferent));
            OnPropertyChanged(nameof(ShippingDifferenceSummary));
            OnPropertyChanged(nameof(IsShippingDifferentFromPayment));
        }

        private static string? GetStringProp(JsonElement root, string name)
        {
            if (root.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String)
                return prop.GetString();
            return null;
        }

        private static string DecodeHtml(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";
            return System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(value)).Trim();
        }

        private static string? NullIfEmpty(string? value) =>
            string.IsNullOrWhiteSpace(value) ? null : value.Trim();

        private static string NormalizeAddressPart(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";
            var decoded = DecodeHtml(value);
            return string.Join(" ", decoded.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                .ToLowerInvariant();
        }
        
        public string Status { get; set; } = "";
        public string PaymentStatus { get; set; } = "Nieznany";
        public string Total { get; set; } = "0.00";
        public string Currency { get; set; } = "PLN";
        public double? CurrencyValue { get; set; }
        public DateTime Date { get; set; }
        
        private string _country = "";
        private string? _isoCode2;
        private string? _isoCode3;
        
        public string Country 
        { 
            get => _country; 
            set 
            { 
                _country = value; 
                OnPropertyChanged(); 
                OnPropertyChanged(nameof(CountryWithIso3));
            } 
        }
        
        public string? IsoCode2 
        { 
            get => _isoCode2; 
            set 
            { 
                _isoCode2 = value; 
                OnPropertyChanged(); 
            } 
        }
        
        public string? IsoCode3 
        { 
            get => _isoCode3; 
            set 
            { 
                _isoCode3 = value; 
                OnPropertyChanged(); 
                OnPropertyChanged(nameof(CountryWithIso3));
            } 
        }
        
        public string AssignedTo { get; set; } = "Nieprzypisane";
        public List<Product> Items { get; set; } = new List<Product>();
        
        public string CountryWithIso3
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Country))
                    return "";
                
                if (!string.IsNullOrWhiteSpace(IsoCode3))
                    return $"{Country} ({IsoCode3})";
                
                return Country;
            }
        }
        public double? CouponAmount { get; set; }
        public double? SubTotal { get; set; }
        public string? CouponTitle { get; set; }
        public double? HandlingAmountNetto { get; set; }
        public double? ShippingAmountNetto { get; set; }
        public double? CodFeeAmountNetto { get; set; }
        public double? GlsAmountNetto { get; set; }
        public double? GlsKgAmountNetto { get; set; }
        public bool UseEuVatRate { get; set; } = false;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isDocumentExists = false;
        public bool IsDocumentExists
        {
            get => _isDocumentExists;
            set
            {
                if (_isDocumentExists != value)
                {
                    _isDocumentExists = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _subiektDocumentNumber = "";
        public string SubiektDocumentNumber
        {
            get => _subiektDocumentNumber;
            set
            {
                if (_subiektDocumentNumber != value)
                {
                    _subiektDocumentNumber = value ?? "";
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DisplayId)); // Powiadom o zmianie DisplayId
                }
            }
        }

        /// <summary>
        /// Zwraca ID zamówienia z numerem ZK, jeśli dokument istnieje
        /// </summary>
        public string DisplayId
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(SubiektDocumentNumber))
                {
                    return $"{Id} ({SubiektDocumentNumber})";
                }
                return Id;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

