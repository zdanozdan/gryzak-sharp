using System.Text.Json;
using System.Text.Json.Serialization;

namespace Gryzak.Models
{
    public class GlsEnvironmentSettings
    {
        public string ApiUrl { get; set; } = "";
        public string UserName { get; set; } = "";
        public string Password { get; set; } = "";
        public int TimeoutSeconds { get; set; } = 30;
    }

    public class GlsConfig
    {
        public const string DefaultTestApiUrl = "https://ade-test.gls-poland.com/adeplus/pm1/ade_webapi2.php?wsdl";
        public const string DefaultProductionApiUrl = "https://adeplus.gls-poland.com/adeplus/pm1/ade_webapi2.php?wsdl";

        public bool UseProduction { get; set; } = false;
        public GlsEnvironmentSettings Test { get; set; } = CreateTestDefaults();
        public GlsEnvironmentSettings Production { get; set; } = CreateProductionDefaults();

        /// <summary>Preferowana drukarka etykiet (np. Zebra).</summary>
        public string LabelPrinterName { get; set; } = "";
        /// <summary>Podgląd etykiety: false = pionowo (domyślnie), true = poziomo.</summary>
        public bool LabelPreviewLandscape { get; set; } = false;
        /// <summary>Czy panel GLS (przygotowalnia, nadania) jest widoczny na liście i w szczegółach dokumentu.</summary>
        public bool GlsPanelEnabled { get; set; } = false;

        [JsonIgnore]
        public GlsEnvironmentSettings Active => UseProduction ? Production : Test;

        [JsonIgnore]
        public string UserName => Active.UserName ?? "";

        [JsonIgnore]
        public string Password => Active.Password ?? "";

        [JsonIgnore]
        public int TimeoutSeconds => Active.TimeoutSeconds;

        public string GetActiveApiUrl()
        {
            var url = Active.ApiUrl;
            return string.IsNullOrWhiteSpace(url)
                ? (UseProduction ? DefaultProductionApiUrl : DefaultTestApiUrl)
                : url.Trim();
        }

        public string GetEnvironmentName()
        {
            return UseProduction ? "live" : "test";
        }

        public void Normalize()
        {
            Test ??= CreateTestDefaults();
            Production ??= CreateProductionDefaults();

            if (string.IsNullOrWhiteSpace(Test.ApiUrl))
            {
                Test.ApiUrl = DefaultTestApiUrl;
            }

            if (string.IsNullOrWhiteSpace(Production.ApiUrl))
            {
                Production.ApiUrl = DefaultProductionApiUrl;
            }

            Test.UserName ??= "";
            Test.Password ??= "";
            Production.UserName ??= "";
            Production.Password ??= "";

            Test.TimeoutSeconds = ClampTimeout(Test.TimeoutSeconds);
            Production.TimeoutSeconds = ClampTimeout(Production.TimeoutSeconds);
        }

        public void ApplyLegacy(JsonElement root)
        {
            Test ??= CreateTestDefaults();
            Production ??= CreateProductionDefaults();

            var legacyUserName = GetLegacyString(root, "UserName");
            var legacyPassword = GetLegacyString(root, "Password");
            var legacyTimeout = GetLegacyInt(root, "TimeoutSeconds");
            var legacyTestUrl = GetLegacyString(root, "TestApiUrl");
            var legacyProductionUrl = GetLegacyString(root, "ProductionApiUrl");

            if (string.IsNullOrWhiteSpace(Test.UserName) && !string.IsNullOrWhiteSpace(legacyUserName))
            {
                Test.UserName = legacyUserName;
            }

            if (string.IsNullOrWhiteSpace(Production.UserName) && !string.IsNullOrWhiteSpace(legacyUserName))
            {
                Production.UserName = legacyUserName;
            }

            if (string.IsNullOrWhiteSpace(Test.Password) && !string.IsNullOrWhiteSpace(legacyPassword))
            {
                Test.Password = legacyPassword;
            }

            if (string.IsNullOrWhiteSpace(Production.Password) && !string.IsNullOrWhiteSpace(legacyPassword))
            {
                Production.Password = legacyPassword;
            }

            if (legacyTimeout is >= 5 and <= 300)
            {
                if (!HasNestedTimeout(root, "Test"))
                {
                    Test.TimeoutSeconds = legacyTimeout.Value;
                }

                if (!HasNestedTimeout(root, "Production"))
                {
                    Production.TimeoutSeconds = legacyTimeout.Value;
                }
            }

            if (string.IsNullOrWhiteSpace(Test.ApiUrl) && !string.IsNullOrWhiteSpace(legacyTestUrl))
            {
                Test.ApiUrl = legacyTestUrl;
            }

            if (string.IsNullOrWhiteSpace(Production.ApiUrl) && !string.IsNullOrWhiteSpace(legacyProductionUrl))
            {
                Production.ApiUrl = legacyProductionUrl;
            }
        }

        private static GlsEnvironmentSettings CreateTestDefaults()
        {
            return new GlsEnvironmentSettings { ApiUrl = DefaultTestApiUrl };
        }

        private static GlsEnvironmentSettings CreateProductionDefaults()
        {
            return new GlsEnvironmentSettings { ApiUrl = DefaultProductionApiUrl };
        }

        private static int ClampTimeout(int timeout)
        {
            return timeout < 5 || timeout > 300 ? 30 : timeout;
        }

        private static string GetLegacyString(JsonElement root, string name)
        {
            return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString() ?? ""
                : "";
        }

        private static int? GetLegacyInt(JsonElement root, string name)
        {
            if (!root.TryGetProperty(name, out var value))
            {
                return null;
            }

            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            {
                return number;
            }

            if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out var parsed))
            {
                return parsed;
            }

            return null;
        }

        private static bool HasNestedTimeout(JsonElement root, string environmentName)
        {
            return root.TryGetProperty(environmentName, out var env)
                && env.ValueKind == JsonValueKind.Object
                && env.TryGetProperty("TimeoutSeconds", out _);
        }
    }
}
