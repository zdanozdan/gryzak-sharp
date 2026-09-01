using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Gryzak.Views
{
    public partial class SelectKontrahentDialog : Window, INotifyPropertyChanged
    {
        private bool _hasOrderContext;

        public KontrahentItem? SelectedKontrahent { get; private set; }
        public bool ShouldAddNew { get; private set; }
        public bool ShouldOpenEmpty { get; private set; }

        public bool HasOrderContext
        {
            get => _hasOrderContext;
            private set { _hasOrderContext = value; OnPropertyChanged(); }
        }

        public SelectKontrahentDialog(
            ObservableCollection<KontrahentItem> kontrahenci,
            string? customerName = null,
            string? email = null,
            string? phone = null,
            string? company = null,
            string? nip = null,
            string? address = null,
            bool showZkActions = true)
        {
            InitializeComponent();
            DataContext = this;

            KontrahenciList.ItemsSource = kontrahenci;
            EmptyListText.Visibility = kontrahenci.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            ZkActionsPanel.Visibility = showZkActions ? Visibility.Visible : Visibility.Collapsed;

            var infoParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(customerName))
                infoParts.Add(customerName);
            if (!string.IsNullOrWhiteSpace(company))
                infoParts.Add(company);
            if (!string.IsNullOrWhiteSpace(email) && !string.Equals(email, "Brak email", StringComparison.OrdinalIgnoreCase))
                infoParts.Add(email);
            if (!string.IsNullOrWhiteSpace(phone) && !string.Equals(phone, "Brak telefonu", StringComparison.OrdinalIgnoreCase))
                infoParts.Add(phone);
            if (!string.IsNullOrWhiteSpace(nip))
                infoParts.Add($"NIP {nip}");
            if (!string.IsNullOrWhiteSpace(address))
                infoParts.Add(address);

            if (infoParts.Count > 0)
            {
                CustomerInfoTextBlock.Text = string.Join(" · ", infoParts);
                HasOrderContext = true;
            }

            SelectPreferredKontrahent(kontrahenci, nip);
        }

        private void SelectPreferredKontrahent(ObservableCollection<KontrahentItem> kontrahenci, string? nip)
        {
            var nipDigits = NipDigits(nip);
            if (!string.IsNullOrEmpty(nipDigits))
            {
                foreach (var item in kontrahenci)
                {
                    if (NipDigitsMatch(item.NIP, nipDigits))
                        item.IsNipMatch = true;
                }
            }

            // Match po NIP zawsze na górze listy.
            if (kontrahenci.Any(k => k.IsNipMatch))
            {
                var ordered = kontrahenci
                    .OrderByDescending(k => k.IsNipMatch)
                    .ThenByDescending(k => k.IsEmailMatch)
                    .ThenByDescending(k => k.IsNameMatch)
                    .ThenByDescending(k => k.IsCompanyMatch)
                    .ToList();
                kontrahenci.Clear();
                foreach (var item in ordered)
                    kontrahenci.Add(item);
            }

            var match = kontrahenci.FirstOrDefault(k => k.IsNipMatch);
            if (match != null)
            {
                KontrahenciList.SelectedItem = match;
                Loaded += (_, _) => KontrahenciList.ScrollIntoView(match);
                return;
            }

            if (kontrahenci.Count == 1)
                KontrahenciList.SelectedIndex = 0;
        }

        private static string NipDigits(string? nip) =>
            new string((nip ?? "").Where(char.IsDigit).ToArray());

        private static bool NipDigitsMatch(string? candidateNip, string orderNipDigits)
        {
            var candidate = NipDigits(candidateNip);
            if (string.IsNullOrEmpty(candidate))
                return false;
            if (candidate == orderNipDigits)
                return true;
            // Prefiks VIES / różna długość — tylko przy sensownej długości NIP (≥8 cyfr).
            if (candidate.Length < 8 || orderNipDigits.Length < 8)
                return false;
            return candidate.EndsWith(orderNipDigits, StringComparison.Ordinal)
                || orderNipDigits.EndsWith(candidate, StringComparison.Ordinal);
        }

        private void KontrahenciList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectButton.IsEnabled = KontrahenciList.SelectedItem is KontrahentItem;
        }

        private void KontrahenciList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (KontrahenciList.SelectedItem is KontrahentItem kontrahent)
                ConfirmSelection(kontrahent);
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (KontrahenciList.SelectedItem is KontrahentItem kontrahent)
                ConfirmSelection(kontrahent);
        }

        private void ConfirmSelection(KontrahentItem kontrahent)
        {
            SelectedKontrahent = kontrahent;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void EmptyButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldOpenEmpty = true;
            DialogResult = true;
            Close();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldAddNew = true;
            DialogResult = false;
            Close();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class KontrahentItem
    {
        public int Id { get; set; }
        public string Symbol { get; set; } = "";
        public string NazwaPelna { get; set; } = "";
        public string Email { get; set; } = "";
        public string NIP { get; set; } = "";
        public string Adres { get; set; } = "";
        public string Miejscowosc { get; set; } = "";
        public string Kod { get; set; } = "";

        public bool IsNipMatch { get; set; }
        public bool IsEmailMatch { get; set; }
        public bool IsNameMatch { get; set; }
        public bool IsCompanyMatch { get; set; }

        public string DisplayTitle =>
            !string.IsNullOrWhiteSpace(NazwaPelna) ? NazwaPelna.Trim()
            : (!string.IsNullOrWhiteSpace(Symbol) ? Symbol.Trim() : $"kh_Id {Id}");

        public string DisplayMeta
        {
            get
            {
                var parts = new List<string>();
                if (!string.IsNullOrWhiteSpace(Symbol))
                    parts.Add(Symbol.Trim());
                if (!string.IsNullOrWhiteSpace(NIP))
                    parts.Add($"NIP {NIP.Trim()}");
                if (!string.IsNullOrWhiteSpace(Email))
                    parts.Add(Email.Trim());
                return string.Join(" · ", parts);
            }
        }

        public string DisplayAddress
        {
            get
            {
                var cityLine = $"{Kod} {Miejscowosc}".Trim();
                if (!string.IsNullOrWhiteSpace(Adres) && !string.IsNullOrWhiteSpace(cityLine))
                    return $"{Adres.Trim()}, {cityLine}";
                if (!string.IsNullOrWhiteSpace(Adres))
                    return Adres.Trim();
                return cityLine;
            }
        }
    }
}
