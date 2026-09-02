using System.Text.Json;
using System.Text.Json.Serialization;

namespace Gryzak.Models
{
    public class ShopEnvironmentSettings
    {
        public string ApiUrl { get; set; } = "";
        public string ApiToken { get; set; } = "";
        public int ApiTimeout { get; set; } = 30;
        public string OrderListEndpoint { get; set; } = "/index.php?route=extension/module/orders/list&format=json";
        public string OrderDetailsEndpoint { get; set; } = "/index.php?route=extension/module/orders/details&token=strefalicencji&order_id={order_id}&format=json";
    }

    public class ApiConfig
    {
        public ShopEnvironmentSettings Test { get; set; } = CreateDefaults();
        public ShopEnvironmentSettings Production { get; set; } = CreateDefaults();

        /// <summary>Ostatnio wybrana zakładka główna (0 = zamówienia, 1 = Nadania GLS).</summary>
        public int SelectedMainTabIndex { get; set; }

        /// <summary>Ustawiane przez ConfigService z globalnego environment.json — nie serializować tu.</summary>
        [JsonIgnore]
        public bool UseProduction { get; set; }

        [JsonIgnore]
        public ShopEnvironmentSettings Active => UseProduction ? Production : Test;

        [JsonIgnore]
        public string ApiUrl
        {
            get => Active.ApiUrl ?? "";
            set => Active.ApiUrl = value ?? "";
        }

        [JsonIgnore]
        public string ApiToken
        {
            get => Active.ApiToken ?? "";
            set => Active.ApiToken = value ?? "";
        }

        [JsonIgnore]
        public int ApiTimeout
        {
            get => Active.ApiTimeout;
            set => Active.ApiTimeout = value;
        }

        [JsonIgnore]
        public string OrderListEndpoint
        {
            get => Active.OrderListEndpoint ?? "";
            set => Active.OrderListEndpoint = value ?? "";
        }

        [JsonIgnore]
        public string OrderDetailsEndpoint
        {
            get => Active.OrderDetailsEndpoint ?? "";
            set => Active.OrderDetailsEndpoint = value ?? "";
        }

        public void Normalize()
        {
            Test ??= CreateDefaults();
            Production ??= CreateDefaults();

            Test.ApiUrl ??= "";
            Test.ApiToken ??= "";
            Test.OrderListEndpoint ??= CreateDefaults().OrderListEndpoint;
            Test.OrderDetailsEndpoint ??= CreateDefaults().OrderDetailsEndpoint;
            Test.ApiTimeout = ClampTimeout(Test.ApiTimeout);

            Production.ApiUrl ??= "";
            Production.ApiToken ??= "";
            Production.OrderListEndpoint ??= CreateDefaults().OrderListEndpoint;
            Production.OrderDetailsEndpoint ??= CreateDefaults().OrderDetailsEndpoint;
            Production.ApiTimeout = ClampTimeout(Production.ApiTimeout);

            if (string.IsNullOrWhiteSpace(Test.OrderListEndpoint))
            {
                Test.OrderListEndpoint = CreateDefaults().OrderListEndpoint;
            }

            if (string.IsNullOrWhiteSpace(Test.OrderDetailsEndpoint))
            {
                Test.OrderDetailsEndpoint = CreateDefaults().OrderDetailsEndpoint;
            }

            if (string.IsNullOrWhiteSpace(Production.OrderListEndpoint))
            {
                Production.OrderListEndpoint = CreateDefaults().OrderListEndpoint;
            }

            if (string.IsNullOrWhiteSpace(Production.OrderDetailsEndpoint))
            {
                Production.OrderDetailsEndpoint = CreateDefaults().OrderDetailsEndpoint;
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

            // Stary płaski format → kopiuj do obu środowisk (użytkownik rozdzieli później).
            var legacy = ReadLegacyFlat(root);
            Test = Clone(legacy);
            Production = Clone(legacy);
        }

        private static ShopEnvironmentSettings ReadLegacyFlat(JsonElement root)
        {
            var settings = CreateDefaults();
            settings.ApiUrl = GetString(root, "ApiUrl");
            settings.ApiToken = GetString(root, "ApiToken");
            settings.ApiTimeout = GetInt(root, "ApiTimeout") ?? 30;
            var list = GetString(root, "OrderListEndpoint");
            var details = GetString(root, "OrderDetailsEndpoint");
            if (!string.IsNullOrWhiteSpace(list))
            {
                settings.OrderListEndpoint = list;
            }

            if (!string.IsNullOrWhiteSpace(details))
            {
                settings.OrderDetailsEndpoint = details;
            }

            return settings;
        }

        private static ShopEnvironmentSettings CreateDefaults()
        {
            return new ShopEnvironmentSettings();
        }

        private static ShopEnvironmentSettings Clone(ShopEnvironmentSettings source)
        {
            return new ShopEnvironmentSettings
            {
                ApiUrl = source.ApiUrl ?? "",
                ApiToken = source.ApiToken ?? "",
                ApiTimeout = source.ApiTimeout,
                OrderListEndpoint = source.OrderListEndpoint ?? "",
                OrderDetailsEndpoint = source.OrderDetailsEndpoint ?? ""
            };
        }

        private static int ClampTimeout(int timeout) => timeout < 5 || timeout > 300 ? 30 : timeout;

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
