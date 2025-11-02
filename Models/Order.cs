using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

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
        
        public string Id { get => _id; set { _id = value; OnPropertyChanged(); } }
        public string Customer { get => _customer; set { _customer = value; OnPropertyChanged(); } }
        public string Email { get => _email; set { _email = value; OnPropertyChanged(); } }
        public string Phone { get => _phone; set { _phone = value; OnPropertyChanged(); } }
        public string? Company { get => _company; set { _company = value; OnPropertyChanged(); } }
        public string? Nip { get => _nip; set { _nip = value; OnPropertyChanged(); } }
        public string? Address { get => _address; set { _address = value; OnPropertyChanged(); } }
        public string? PaymentAddress1 { get => _paymentAddress1; set { _paymentAddress1 = value; OnPropertyChanged(); } }
        public string? PaymentAddress2 { get => _paymentAddress2; set { _paymentAddress2 = value; OnPropertyChanged(); } }
        public string? PaymentPostcode { get => _paymentPostcode; set { _paymentPostcode = value; OnPropertyChanged(); } }
        public string? PaymentCity { get => _paymentCity; set { _paymentCity = value; OnPropertyChanged(); } }
        
        public string Status { get; set; } = "";
        public string PaymentStatus { get; set; } = "Nieznany";
        public string Total { get; set; } = "0.00";
        public string Currency { get; set; } = "PLN";
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
        public double? HandlingAmount { get; set; }
        public double? ShippingAmount { get; set; }
        public double? CodFeeAmount { get; set; }
        public double? GlsAmount { get; set; }

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

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

