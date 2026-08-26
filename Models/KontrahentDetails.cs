namespace Gryzak.Models
{
    /// <summary>
    /// Podgląd kontrahenta z GET /kontrahenci/{id} (adres podstawowy + opcjonalnie korespondencyjny i dostawy).
    /// </summary>
    public class KontrahentDetails
    {
        public int Id { get; set; }
        public string Symbol { get; set; } = "";
        public string Email { get; set; } = "";
        public string TypNazwa { get; set; } = "";

        public string Nazwa { get; set; } = "";
        public string NazwaPelna { get; set; } = "";
        public string Nip { get; set; } = "";
        public string Adres { get; set; } = "";
        public string Ulica { get; set; } = "";
        public string Kod { get; set; } = "";
        public string Miejscowosc { get; set; } = "";
        public string Poczta { get; set; } = "";
        public string Telefon { get; set; } = "";
        public string Panstwo { get; set; } = "";
        public string Wojewodztwo { get; set; } = "";

        public bool HasAdresKorespondencyjny { get; set; }
        public string KorespondencyjnyNazwa { get; set; } = "";
        public string KorespondencyjnyNazwaPelna { get; set; } = "";
        public string KorespondencyjnyNip { get; set; } = "";
        public string KorespondencyjnyAdres { get; set; } = "";
        public string KorespondencyjnyUlica { get; set; } = "";
        public string KorespondencyjnyKod { get; set; } = "";
        public string KorespondencyjnyMiejscowosc { get; set; } = "";
        public string KorespondencyjnyPoczta { get; set; } = "";
        public string KorespondencyjnyTelefon { get; set; } = "";
        public string KorespondencyjnyPanstwo { get; set; } = "";
        public string KorespondencyjnyWojewodztwo { get; set; } = "";

        public bool HasAdresDostawy { get; set; }
        public string DostawaNazwa { get; set; } = "";
        public string DostawaNazwaPelna { get; set; } = "";
        public string DostawaNip { get; set; } = "";
        public string DostawaAdres { get; set; } = "";
        public string DostawaUlica { get; set; } = "";
        public string DostawaKod { get; set; } = "";
        public string DostawaMiejscowosc { get; set; } = "";
        public string DostawaPoczta { get; set; } = "";
        public string DostawaTelefon { get; set; } = "";
        public string DostawaPanstwo { get; set; } = "";
        public string DostawaWojewodztwo { get; set; } = "";
    }
}
