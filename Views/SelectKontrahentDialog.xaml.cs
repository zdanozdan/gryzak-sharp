using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace Gryzak.Views
{
    public partial class SelectKontrahentDialog : Window
    {
        public KontrahentItem? SelectedKontrahent { get; private set; }
        public bool ShouldAddNew { get; private set; } = false;
        public bool ShouldOpenEmpty { get; private set; } = false;
        
        public SelectKontrahentDialog(ObservableCollection<KontrahentItem> kontrahenci, string? customerName = null, string? email = null, string? phone = null, string? company = null, string? nip = null, string? address = null)
        {
            InitializeComponent();
            KontrahenciDataGrid.ItemsSource = kontrahenci;
            
            // Ustaw dane kontrahenta z API w nagłówku
            if (CustomerInfoTextBlock != null)
            {
                var infoParts = new System.Collections.Generic.List<string>();
                
                if (!string.IsNullOrWhiteSpace(customerName))
                {
                    infoParts.Add($"Imię i nazwisko: {customerName}");
                }
                
                if (!string.IsNullOrWhiteSpace(company))
                {
                    infoParts.Add($"Firma: {company}");
                }
                
                if (!string.IsNullOrWhiteSpace(email) && email != "Brak email")
                {
                    infoParts.Add($"Email: {email}");
                }
                
                if (!string.IsNullOrWhiteSpace(phone) && phone != "Brak telefonu")
                {
                    infoParts.Add($"Telefon: {phone}");
                }
                
                if (!string.IsNullOrWhiteSpace(nip))
                {
                    infoParts.Add($"NIP: {nip}");
                }
                
                if (!string.IsNullOrWhiteSpace(address))
                {
                    infoParts.Add($"Adres: {address}");
                }
                
                if (infoParts.Count > 0)
                {
                    CustomerInfoTextBlock.Text = string.Join(" | ", infoParts);
                    CustomerInfoTextBlock.Visibility = Visibility.Visible;
                }
                else
                {
                    CustomerInfoTextBlock.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void KontrahenciDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectButton.IsEnabled = KontrahenciDataGrid.SelectedItem != null;
        }

        private void KontrahenciDataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (KontrahenciDataGrid.SelectedItem is KontrahentItem kontrahent)
            {
                SelectedKontrahent = kontrahent;
                DialogResult = true;
                Close();
            }
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (KontrahenciDataGrid.SelectedItem is KontrahentItem kontrahent)
            {
                SelectedKontrahent = kontrahent;
                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Po prostu zamknij okno bez żadnych dodatkowych akcji
            DialogResult = false;
            Close();
        }
        
        private void EmptyButton_Click(object sender, RoutedEventArgs e)
        {
            // Ustaw flagę, że użytkownik chce otworzyć ZK bez kontrahenta
            ShouldOpenEmpty = true;
            DialogResult = true; // true aby wskazać, że wykonujemy akcję (nie anulujemy)
            Close();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            // Ustaw flagę, że użytkownik chce dodać nowego kontrahenta
            ShouldAddNew = true;
            // Zamknij dialog - kod wywołujący dialog sprawdzi ShouldAddNew
            DialogResult = false;
            Close();
        }
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
    }
}

