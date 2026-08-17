namespace Gryzak.Models
{
    public class GlsConfig
    {
        public const string DefaultTestApiUrl = "https://ade-test.gls-poland.com/adeplus/pm1/ade_webapi2.php?wsdl";
        public const string DefaultProductionApiUrl = "https://adeplus.gls-poland.com/adeplus/pm1/ade_webapi2.php?wsdl";

        public string UserName { get; set; } = "";
        public string Password { get; set; } = "";
        public bool UseProduction { get; set; } = false;
        public string TestApiUrl { get; set; } = DefaultTestApiUrl;
        public string ProductionApiUrl { get; set; } = DefaultProductionApiUrl;
        public int TimeoutSeconds { get; set; } = 30;

        public string GetActiveApiUrl()
        {
            var url = UseProduction ? ProductionApiUrl : TestApiUrl;
            return string.IsNullOrWhiteSpace(url)
                ? (UseProduction ? DefaultProductionApiUrl : DefaultTestApiUrl)
                : url.Trim();
        }

        public string GetEnvironmentName()
        {
            return UseProduction ? "produkcja" : "test";
        }
    }
}
