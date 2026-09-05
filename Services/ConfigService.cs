using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using Gryzak.Models;

namespace Gryzak.Services
{
    public class ConfigService
    {
        private readonly string _configPath;
        private readonly string _subiektConfigPath;
        private readonly string _glsConfigPath;
        private readonly string _environmentPath;
        private readonly string _historyPath;

        public ConfigService()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var gryzakPath = Path.Combine(appDataPath, "Gryzak");

            if (!Directory.Exists(gryzakPath))
            {
                Directory.CreateDirectory(gryzakPath);
            }

            _configPath = Path.Combine(gryzakPath, "config.json");
            _subiektConfigPath = Path.Combine(gryzakPath, "subiekt_config.json");
            _glsConfigPath = Path.Combine(gryzakPath, "gls_config.json");
            _environmentPath = Path.Combine(gryzakPath, "environment.json");
            _historyPath = Path.Combine(gryzakPath, "order_history.json");
        }

        public AppEnvironmentConfig LoadEnvironment()
        {
            try
            {
                if (File.Exists(_environmentPath))
                {
                    var json = File.ReadAllText(_environmentPath);
                    var config = JsonSerializer.Deserialize<AppEnvironmentConfig>(json);
                    if (config != null)
                    {
                        return config;
                    }
                }
                else
                {
                    // Migracja: przejmij UseProduction z GLS, jeśli było zapisane.
                    var migrated = TryMigrateEnvironmentFromGls();
                    SaveEnvironment(migrated);
                    return migrated;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd ładowania środowiska: {ex.Message}");
            }

            return new AppEnvironmentConfig();
        }

        public void SaveEnvironment(AppEnvironmentConfig config)
        {
            try
            {
                var json = JsonSerializer.Serialize(config, JsonWriteOptions);
                File.WriteAllText(_environmentPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd zapisywania środowiska: {ex.Message}");
                throw;
            }
        }

        public bool GetUseProduction() => LoadEnvironment().UseProduction;

        public void SetUseProduction(bool useProduction)
        {
            var env = LoadEnvironment();
            if (env.UseProduction == useProduction)
            {
                return;
            }

            env.UseProduction = useProduction;
            SaveEnvironment(env);

            // Utrzymaj zgodność pliku GLS (Active/URL w UI GLS).
            try
            {
                var gls = LoadGlsConfigRaw();
                gls.UseProduction = useProduction;
                SaveGlsConfigRaw(gls);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Nie udało się zsynchronizować GLS UseProduction: {ex.Message}");
            }
        }

        public string GetEnvironmentName() => LoadEnvironment().GetEnvironmentName();

        public string GetEnvironmentDisplayName() => LoadEnvironment().GetEnvironmentDisplayName();

        public ApiConfig LoadConfig()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var json = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<ApiConfig>(json);
                    if (config != null)
                    {
                        using var doc = JsonDocument.Parse(json);
                        config.ApplyLegacy(doc.RootElement);
                        config.Normalize();
                        ApplyActiveEnvironment(config);

                        if (config.OrderDetailsEndpoint.Contains("{order_{d}", StringComparison.Ordinal))
                        {
                            config.OrderDetailsEndpoint =
                                "/index.php?route=extension/module/orders/details&token=strefalicencji&order_id={order_id}&format=json";
                            SaveConfig(config);
                        }

                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd ładowania konfiguracji: {ex.Message}");
            }

            return GetDefaultConfig();
        }

        public void SaveConfig(ApiConfig config)
        {
            try
            {
                config.Normalize();
                var json = JsonSerializer.Serialize(config, JsonWriteOptions);
                File.WriteAllText(_configPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd zapisywania konfiguracji: {ex.Message}");
                throw;
            }
        }

        public void ResetConfig()
        {
            SaveConfig(GetDefaultConfig());
        }

        private ApiConfig GetDefaultConfig()
        {
            var config = new ApiConfig();
            config.Test.ApiUrl = "https://mikran.pl";
            config.Production.ApiUrl = "https://mikran.pl";
            config.Normalize();
            ApplyActiveEnvironment(config);
            return config;
        }

        public SubiektConfig LoadSubiektConfig()
        {
            try
            {
                if (File.Exists(_subiektConfigPath))
                {
                    var json = File.ReadAllText(_subiektConfigPath);
                    var config = JsonSerializer.Deserialize<SubiektConfig>(json);
                    if (config != null)
                    {
                        using var doc = JsonDocument.Parse(json);
                        config.ApplyLegacy(doc.RootElement);
                        config.Normalize();
                        ApplyActiveEnvironment(config);
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd ładowania konfiguracji Subiekt: {ex.Message}");
            }

            return GetDefaultSubiektConfig();
        }

        public void SaveSubiektConfig(SubiektConfig config)
        {
            try
            {
                config.Normalize();
                var json = JsonSerializer.Serialize(config, JsonWriteOptions);
                File.WriteAllText(_subiektConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd zapisywania konfiguracji Subiekt: {ex.Message}");
                throw;
            }
        }

        private SubiektConfig GetDefaultSubiektConfig()
        {
            var config = new SubiektConfig
            {
                AutoReleaseLicenseTimeoutMinutes = 0,
                DiscountCalculationMode = "percent"
            };
            config.Test.ServerAddress = "192.168.0.140";
            config.Test.ServerUsername = "mikran_com";
            config.Test.ServerPassword = "mikran_comqwer4321";
            config.Test.User = "Szef";
            config.Test.Password = "zdanoszef123";
            config.Production = new SubiektEnvironmentSettings
            {
                ServerAddress = config.Test.ServerAddress,
                ServerUsername = config.Test.ServerUsername,
                ServerPassword = config.Test.ServerPassword,
                User = config.Test.User,
                Password = config.Test.Password
            };
            config.Normalize();
            ApplyActiveEnvironment(config);
            return config;
        }

        public GlsConfig LoadGlsConfig()
        {
            var config = LoadGlsConfigRaw();
            config.UseProduction = GetUseProduction();
            return config;
        }

        public void SaveGlsConfig(GlsConfig config)
        {
            config.UseProduction = GetUseProduction();
            SaveGlsConfigRaw(config);
        }

        private GlsConfig LoadGlsConfigRaw()
        {
            try
            {
                if (File.Exists(_glsConfigPath))
                {
                    var json = File.ReadAllText(_glsConfigPath);
                    var config = JsonSerializer.Deserialize<GlsConfig>(json);
                    if (config != null)
                    {
                        using var doc = JsonDocument.Parse(json);
                        config.ApplyLegacy(doc.RootElement);
                        config.Normalize();
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd ładowania konfiguracji GLS: {ex.Message}");
            }

            return GetDefaultGlsConfig();
        }

        private void SaveGlsConfigRaw(GlsConfig config)
        {
            try
            {
                config.Normalize();
                var json = JsonSerializer.Serialize(config, JsonWriteOptions);
                File.WriteAllText(_glsConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd zapisywania konfiguracji GLS: {ex.Message}");
                throw;
            }
        }

        private GlsConfig GetDefaultGlsConfig()
        {
            var config = new GlsConfig();
            config.Normalize();
            return config;
        }

        public List<string> LoadOrderHistory()
        {
            try
            {
                if (File.Exists(_historyPath))
                {
                    var json = File.ReadAllText(_historyPath);
                    var historyData = JsonSerializer.Deserialize<OrderHistoryData>(json);
                    if (historyData?.OrderIds != null && historyData.OrderIds.Count > 0)
                    {
                        return historyData.OrderIds.Take(10).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd ładowania historii zamówień: {ex.Message}");
            }

            return new List<string>();
        }

        public void SaveOrderHistory(List<string> history)
        {
            try
            {
                var historyData = new OrderHistoryData
                {
                    OrderIds = history.Take(10).ToList()
                };
                var json = JsonSerializer.Serialize(historyData, JsonWriteOptions);
                File.WriteAllText(_historyPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd zapisywania historii zamówień: {ex.Message}");
            }
        }

        public void AddOrderToHistory(string orderId)
        {
            if (string.IsNullOrWhiteSpace(orderId))
            {
                return;
            }

            try
            {
                var history = LoadOrderHistory();
                history.RemoveAll(id => id == orderId);
                history.Insert(0, orderId);
                if (history.Count > 10)
                {
                    history = history.Take(10).ToList();
                }

                SaveOrderHistory(history);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd dodawania zamówienia do historii: {ex.Message}");
            }
        }

        private class OrderHistoryData
        {
            public List<string> OrderIds { get; set; } = new List<string>();
        }

        private static JsonSerializerOptions JsonWriteOptions => new() { WriteIndented = true };

        public GryzakSettingsExport CreateSettingsExport()
        {
            return new GryzakSettingsExport
            {
                Version = GryzakSettingsExport.CurrentVersion,
                ExportedAt = DateTime.UtcNow,
                Environment = LoadEnvironment(),
                Shop = LoadConfig(),
                Subiekt = LoadSubiektConfig(),
                Gls = LoadGlsConfig()
            };
        }

        public void ExportSettings(string filePath)
        {
            var export = CreateSettingsExport();
            ObfuscateSecrets(export);
            var json = JsonSerializer.Serialize(export, JsonWriteOptions);
            // Operator Subiekta (User/Password + launch) jest per-stanowisko — nie eksportuj.
            json = RemoveSubiektOperatorSettingsFromExportJson(json);
            File.WriteAllText(filePath, json);
        }

        public void ImportSettings(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Nie znaleziono pliku ustawień.", filePath);
            }

            var localSubiekt = LoadSubiektConfig();

            var json = File.ReadAllText(filePath);
            var export = JsonSerializer.Deserialize<GryzakSettingsExport>(json);
            if (export?.Shop == null || export.Subiekt == null || export.Gls == null)
            {
                throw new InvalidOperationException("Plik nie zawiera kompletnych ustawień Gryzak (sklep, Subiekt, GLS).");
            }

            using (var doc = JsonDocument.Parse(json))
            {
                if (doc.RootElement.TryGetProperty("Shop", out var shopEl))
                {
                    export.Shop.ApplyLegacy(shopEl);
                }

                if (doc.RootElement.TryGetProperty("Subiekt", out var subiektEl))
                {
                    export.Subiekt.ApplyLegacy(subiektEl);
                }

                if (doc.RootElement.TryGetProperty("Gls", out var glsEl))
                {
                    export.Gls.ApplyLegacy(glsEl);
                }

                // Operator Subiekta jest per-stanowisko — nigdy nie nadpisuj z pliku eksportu.
                PreserveLocalSubiektOperatorSettings(export.Subiekt, localSubiekt);
            }

            DeobfuscateSecrets(export);

            export.Shop.Normalize();
            export.Subiekt.Normalize();
            export.Gls.Normalize();

            if (export.Environment != null)
            {
                SaveEnvironment(export.Environment);
            }
            else
            {
                // Stary eksport: weź UseProduction z GLS jeśli jest.
                SaveEnvironment(new AppEnvironmentConfig { UseProduction = export.Gls.UseProduction });
            }

            SaveConfig(export.Shop);
            SaveSubiektConfig(export.Subiekt);
            SaveGlsConfig(export.Gls);
        }

        /// <summary>
        /// Pola operatora / uruchomienia Sfery — nie przenosimy między stacjami.
        /// </summary>
        private static readonly string[] SubiektOperatorExportPropertyNames =
        {
            "User",
            "Password",
            "GtProdukt",
            "AuthenticationMode",
            "LaunchDopasujOperatora",
            "LaunchTryb"
        };

        private static string RemoveSubiektOperatorSettingsFromExportJson(string json)
        {
            var root = JsonNode.Parse(json) as JsonObject;
            if (root == null)
                return json;

            if (root["Subiekt"] is JsonObject subiekt)
            {
                RemoveOperatorProps(subiekt["Test"] as JsonObject);
                RemoveOperatorProps(subiekt["Production"] as JsonObject);
            }

            return root.ToJsonString(JsonWriteOptions);
        }

        private static void RemoveOperatorProps(JsonObject? env)
        {
            if (env == null)
                return;

            foreach (var name in SubiektOperatorExportPropertyNames)
                env.Remove(name);
        }

        private static void PreserveLocalSubiektOperatorSettings(SubiektConfig imported, SubiektConfig local)
        {
            imported.Test ??= new SubiektEnvironmentSettings();
            imported.Production ??= new SubiektEnvironmentSettings();
            local.Test ??= new SubiektEnvironmentSettings();
            local.Production ??= new SubiektEnvironmentSettings();

            CopyOperatorEnv(local.Test, imported.Test);
            CopyOperatorEnv(local.Production, imported.Production);
        }

        private static void CopyOperatorEnv(SubiektEnvironmentSettings from, SubiektEnvironmentSettings to)
        {
            to.User = from.User ?? "";
            to.Password = from.Password ?? "";
            to.GtProdukt = from.GtProdukt;
            to.AuthenticationMode = from.AuthenticationMode;
            to.LaunchDopasujOperatora = from.LaunchDopasujOperatora;
            to.LaunchTryb = from.LaunchTryb;
        }

        private static void ObfuscateSecrets(GryzakSettingsExport export)
        {
            if (export.Shop != null)
            {
                export.Shop.Test ??= new ShopEnvironmentSettings();
                export.Shop.Production ??= new ShopEnvironmentSettings();
                export.Shop.Test.ApiToken = SettingsSecretCodec.Encode(export.Shop.Test.ApiToken);
                export.Shop.Production.ApiToken = SettingsSecretCodec.Encode(export.Shop.Production.ApiToken);
            }

            if (export.Subiekt != null)
            {
                export.Subiekt.Test ??= new SubiektEnvironmentSettings();
                export.Subiekt.Production ??= new SubiektEnvironmentSettings();
                ObfuscateSubiektEnv(export.Subiekt.Test);
                ObfuscateSubiektEnv(export.Subiekt.Production);
            }

            if (export.Gls != null)
            {
                export.Gls.Test ??= new GlsEnvironmentSettings();
                export.Gls.Production ??= new GlsEnvironmentSettings();
                export.Gls.Test.Password = SettingsSecretCodec.Encode(export.Gls.Test.Password);
                export.Gls.Production.Password = SettingsSecretCodec.Encode(export.Gls.Production.Password);
            }
        }

        private static void DeobfuscateSecrets(GryzakSettingsExport export)
        {
            if (export.Shop != null)
            {
                export.Shop.Test ??= new ShopEnvironmentSettings();
                export.Shop.Production ??= new ShopEnvironmentSettings();
                export.Shop.Test.ApiToken = SettingsSecretCodec.Decode(export.Shop.Test.ApiToken);
                export.Shop.Production.ApiToken = SettingsSecretCodec.Decode(export.Shop.Production.ApiToken);
            }

            if (export.Subiekt != null)
            {
                export.Subiekt.Test ??= new SubiektEnvironmentSettings();
                export.Subiekt.Production ??= new SubiektEnvironmentSettings();
                DeobfuscateSubiektEnv(export.Subiekt.Test);
                DeobfuscateSubiektEnv(export.Subiekt.Production);
            }

            if (export.Gls != null)
            {
                export.Gls.Test ??= new GlsEnvironmentSettings();
                export.Gls.Production ??= new GlsEnvironmentSettings();
                export.Gls.Test.Password = SettingsSecretCodec.Decode(export.Gls.Test.Password);
                export.Gls.Production.Password = SettingsSecretCodec.Decode(export.Gls.Production.Password);
            }
        }

        private static void ObfuscateSubiektEnv(SubiektEnvironmentSettings env)
        {
            env.ApiKey = SettingsSecretCodec.Encode(env.ApiKey);
            env.ServerPassword = SettingsSecretCodec.Encode(env.ServerPassword);
        }

        private static void DeobfuscateSubiektEnv(SubiektEnvironmentSettings env)
        {
            env.ApiKey = SettingsSecretCodec.Decode(env.ApiKey);
            env.ServerPassword = SettingsSecretCodec.Decode(env.ServerPassword);
            env.Password = SettingsSecretCodec.Decode(env.Password);
        }

        private void ApplyActiveEnvironment(ApiConfig config)
        {
            config.UseProduction = GetUseProduction();
        }

        private void ApplyActiveEnvironment(SubiektConfig config)
        {
            config.UseProduction = GetUseProduction();
        }

        private AppEnvironmentConfig TryMigrateEnvironmentFromGls()
        {
            try
            {
                if (File.Exists(_glsConfigPath))
                {
                    var json = File.ReadAllText(_glsConfigPath);
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("UseProduction", out var value)
                        && (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False))
                    {
                        return new AppEnvironmentConfig { UseProduction = value.GetBoolean() };
                    }
                }
            }
            catch
            {
                // ignore
            }

            return new AppEnvironmentConfig();
        }
    }
}
