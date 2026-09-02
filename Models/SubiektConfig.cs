using System.Text.Json;
using System.Text.Json.Serialization;

namespace Gryzak.Models
{
    public class SubiektEnvironmentSettings
    {
        /// <summary>Base URL Subiekt REST API, np. http://192.168.0.140:5082/api/v1</summary>
        public string ApiBaseUrl { get; set; } = "";

        /// <summary>Opcjonalny klucz API (nagłówek X-Api-Key).</summary>
        public string ApiKey { get; set; } = "";

        public string ServerAddress { get; set; } = "";
        public string DatabaseName { get; set; } = "";
        public string ServerUsername { get; set; } = "";
        public string ServerPassword { get; set; } = "";
        public string User { get; set; } = "Szef";
        public string Password { get; set; } = "";
        public int GtProdukt { get; set; } = 1;
        public int AuthenticationMode { get; set; }
        public int LaunchDopasujOperatora { get; set; } = 2;
        public int LaunchTryb { get; set; }
    }

    public class SubiektConfig
    {
        public SubiektEnvironmentSettings Test { get; set; } = CreateDefaults();
        public SubiektEnvironmentSettings Production { get; set; } = CreateDefaults();

        public int AutoReleaseLicenseTimeoutMinutes { get; set; }
        public string DiscountCalculationMode { get; set; } = "percent";
        public bool CalculateFromGrossPrices { get; set; }
        public string DiscountRoundingMode { get; set; } = "percent";

        [JsonIgnore]
        public bool UseProduction { get; set; }

        [JsonIgnore]
        public SubiektEnvironmentSettings Active => UseProduction ? Production : Test;

        [JsonIgnore]
        public string ApiBaseUrl
        {
            get => Active.ApiBaseUrl ?? "";
            set => Active.ApiBaseUrl = value ?? "";
        }

        [JsonIgnore]
        public string ApiKey
        {
            get => Active.ApiKey ?? "";
            set => Active.ApiKey = value ?? "";
        }

        [JsonIgnore]
        public string ServerAddress
        {
            get => Active.ServerAddress ?? "";
            set => Active.ServerAddress = value ?? "";
        }

        [JsonIgnore]
        public string DatabaseName
        {
            get => Active.DatabaseName ?? "";
            set => Active.DatabaseName = value ?? "";
        }

        [JsonIgnore]
        public string ServerUsername
        {
            get => Active.ServerUsername ?? "";
            set => Active.ServerUsername = value ?? "";
        }

        [JsonIgnore]
        public string ServerPassword
        {
            get => Active.ServerPassword ?? "";
            set => Active.ServerPassword = value ?? "";
        }

        [JsonIgnore]
        public string User
        {
            get => Active.User ?? "";
            set => Active.User = value ?? "";
        }

        [JsonIgnore]
        public string Password
        {
            get => Active.Password ?? "";
            set => Active.Password = value ?? "";
        }

        [JsonIgnore]
        public int GtProdukt
        {
            get => Active.GtProdukt;
            set => Active.GtProdukt = value;
        }

        [JsonIgnore]
        public int AuthenticationMode
        {
            get => Active.AuthenticationMode;
            set => Active.AuthenticationMode = value;
        }

        [JsonIgnore]
        public int LaunchDopasujOperatora
        {
            get => Active.LaunchDopasujOperatora;
            set => Active.LaunchDopasujOperatora = value;
        }

        [JsonIgnore]
        public int LaunchTryb
        {
            get => Active.LaunchTryb;
            set => Active.LaunchTryb = value;
        }

        public void Normalize()
        {
            Test ??= CreateDefaults();
            Production ??= CreateDefaults();
            NormalizeEnv(Test);
            NormalizeEnv(Production);

            if (string.IsNullOrWhiteSpace(DiscountCalculationMode))
            {
                DiscountCalculationMode = "percent";
            }

            if (string.IsNullOrWhiteSpace(DiscountRoundingMode))
            {
                DiscountRoundingMode = "percent";
            }

            if (AutoReleaseLicenseTimeoutMinutes < 0)
            {
                AutoReleaseLicenseTimeoutMinutes = 0;
            }
        }

        public void ApplyLegacy(JsonElement root)
        {
            Test ??= CreateDefaults();
            Production ??= CreateDefaults();

            if (root.TryGetProperty("Test", out var testEl) && testEl.ValueKind == JsonValueKind.Object)
            {
                return;
            }

            var legacy = ReadLegacyFlat(root);
            Test = Clone(legacy);
            Production = Clone(legacy);
        }

        private static void NormalizeEnv(SubiektEnvironmentSettings env)
        {
            env.ApiBaseUrl ??= "";
            env.ApiKey ??= "";
            env.ServerAddress ??= "";
            env.DatabaseName ??= "";
            env.ServerUsername ??= "";
            env.ServerPassword ??= "";
            env.User ??= "Szef";
            env.Password ??= "";
        }

        private static SubiektEnvironmentSettings CreateDefaults()
        {
            return new SubiektEnvironmentSettings();
        }

        private static SubiektEnvironmentSettings Clone(SubiektEnvironmentSettings source)
        {
            return new SubiektEnvironmentSettings
            {
                ApiBaseUrl = source.ApiBaseUrl ?? "",
                ApiKey = source.ApiKey ?? "",
                ServerAddress = source.ServerAddress ?? "",
                DatabaseName = source.DatabaseName ?? "",
                ServerUsername = source.ServerUsername ?? "",
                ServerPassword = source.ServerPassword ?? "",
                User = source.User ?? "Szef",
                Password = source.Password ?? "",
                GtProdukt = source.GtProdukt,
                AuthenticationMode = source.AuthenticationMode,
                LaunchDopasujOperatora = source.LaunchDopasujOperatora,
                LaunchTryb = source.LaunchTryb
            };
        }

        private static SubiektEnvironmentSettings ReadLegacyFlat(JsonElement root)
        {
            var settings = CreateDefaults();
            settings.ApiBaseUrl = GetString(root, "ApiBaseUrl");
            settings.ApiKey = GetString(root, "ApiKey");
            settings.ServerAddress = GetString(root, "ServerAddress");
            settings.DatabaseName = GetString(root, "DatabaseName");
            settings.ServerUsername = GetString(root, "ServerUsername");
            settings.ServerPassword = GetString(root, "ServerPassword");
            var user = GetString(root, "User");
            if (!string.IsNullOrWhiteSpace(user))
            {
                settings.User = user;
            }

            settings.Password = GetString(root, "Password");
            settings.GtProdukt = GetInt(root, "GtProdukt") ?? 1;
            settings.AuthenticationMode = GetInt(root, "AuthenticationMode") ?? 0;
            settings.LaunchDopasujOperatora = GetInt(root, "LaunchDopasujOperatora") ?? 2;
            settings.LaunchTryb = GetInt(root, "LaunchTryb") ?? 0;
            return settings;
        }

        private static string GetString(JsonElement root, string name) =>
            root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString() ?? ""
                : "";

        private static int? GetInt(JsonElement root, string name)
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
    }
}
