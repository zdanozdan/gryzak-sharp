namespace Gryzak.Models
{
    /// <summary>Globalny przełącznik Test / Live dla całej aplikacji.</summary>
    public class AppEnvironmentConfig
    {
        public bool UseProduction { get; set; }

        public string GetEnvironmentName() => UseProduction ? "live" : "test";

        public string GetEnvironmentDisplayName() => UseProduction ? "Live" : "TEST";
    }
}
