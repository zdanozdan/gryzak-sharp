using System.Windows;
using Gryzak.Models;

namespace Gryzak.Views
{
    public partial class KontrahentDetailsDialog : Window
    {
        public KontrahentDetailsDialog(KontrahentDetails details, string? roleTitle = null)
        {
            InitializeComponent();
            DataContext = new ViewModel(details, roleTitle);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private sealed class ViewModel
        {
            public ViewModel(KontrahentDetails d, string? roleTitle)
            {
                var title = string.IsNullOrWhiteSpace(roleTitle) ? "Kontrahent" : roleTitle.Trim();
                var name = !string.IsNullOrWhiteSpace(d.NazwaPelna) ? d.NazwaPelna
                    : (!string.IsNullOrWhiteSpace(d.Nazwa) ? d.Nazwa : d.Symbol);
                HeaderTitle = string.IsNullOrWhiteSpace(name) ? title : $"{title}: {name}";
                HeaderSubtitle = d.Id > 0
                    ? $"kh_Id = {d.Id}" + (string.IsNullOrWhiteSpace(d.Symbol) ? "" : $"  ·  {d.Symbol}")
                    : "";

                IdDisplay = d.Id > 0 ? d.Id.ToString() : "";
                Symbol = d.Symbol;
                TypNazwa = d.TypNazwa;
                Email = d.Email;

                Nazwa = d.Nazwa;
                NazwaPelna = d.NazwaPelna;
                Nip = d.Nip;
                Adres = d.Adres;
                Ulica = d.Ulica;
                Kod = d.Kod;
                Miejscowosc = d.Miejscowosc;
                Poczta = d.Poczta;
                Telefon = d.Telefon;
                Panstwo = d.Panstwo;
                Wojewodztwo = d.Wojewodztwo;

                HasAdresKorespondencyjny = d.HasAdresKorespondencyjny;
                KorespondencyjnyNazwa = d.KorespondencyjnyNazwa;
                KorespondencyjnyNazwaPelna = d.KorespondencyjnyNazwaPelna;
                KorespondencyjnyNip = d.KorespondencyjnyNip;
                KorespondencyjnyAdres = d.KorespondencyjnyAdres;
                KorespondencyjnyUlica = d.KorespondencyjnyUlica;
                KorespondencyjnyKod = d.KorespondencyjnyKod;
                KorespondencyjnyMiejscowosc = d.KorespondencyjnyMiejscowosc;
                KorespondencyjnyTelefon = d.KorespondencyjnyTelefon;
                KorespondencyjnyPanstwo = d.KorespondencyjnyPanstwo;

                HasAdresDostawy = d.HasAdresDostawy;
                DostawaNazwa = d.DostawaNazwa;
                DostawaNazwaPelna = d.DostawaNazwaPelna;
                DostawaNip = d.DostawaNip;
                DostawaAdres = d.DostawaAdres;
                DostawaUlica = d.DostawaUlica;
                DostawaKod = d.DostawaKod;
                DostawaMiejscowosc = d.DostawaMiejscowosc;
                DostawaTelefon = d.DostawaTelefon;
                DostawaPanstwo = d.DostawaPanstwo;
            }

            public string HeaderTitle { get; }
            public string HeaderSubtitle { get; }
            public string IdDisplay { get; }
            public string Symbol { get; }
            public string TypNazwa { get; }
            public string Email { get; }
            public string Nazwa { get; }
            public string NazwaPelna { get; }
            public string Nip { get; }
            public string Adres { get; }
            public string Ulica { get; }
            public string Kod { get; }
            public string Miejscowosc { get; }
            public string Poczta { get; }
            public string Telefon { get; }
            public string Panstwo { get; }
            public string Wojewodztwo { get; }
            public bool HasAdresKorespondencyjny { get; }
            public string KorespondencyjnyNazwa { get; }
            public string KorespondencyjnyNazwaPelna { get; }
            public string KorespondencyjnyNip { get; }
            public string KorespondencyjnyAdres { get; }
            public string KorespondencyjnyUlica { get; }
            public string KorespondencyjnyKod { get; }
            public string KorespondencyjnyMiejscowosc { get; }
            public string KorespondencyjnyTelefon { get; }
            public string KorespondencyjnyPanstwo { get; }
            public bool HasAdresDostawy { get; }
            public string DostawaNazwa { get; }
            public string DostawaNazwaPelna { get; }
            public string DostawaNip { get; }
            public string DostawaAdres { get; }
            public string DostawaUlica { get; }
            public string DostawaKod { get; }
            public string DostawaMiejscowosc { get; }
            public string DostawaTelefon { get; }
            public string DostawaPanstwo { get; }
        }
    }
}
