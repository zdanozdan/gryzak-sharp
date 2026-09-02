namespace Gryzak.Models
{
    /// <summary>Globalny przełącznik Test / Produkcja dla całej aplikacji.</summary>
    public class AppEnvironmentConfig
    {
        public bool UseProduction { get; set; }

        public string GetEnvironmentName() => UseProduction ? "produkcja" : "test";

        public string GetEnvironmentDisplayName() => UseProduction ? "Live" : "TEST";
    }
}
