using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Gryzak.Models;

namespace Gryzak.Services
{
    public class ConfigService
    {
        private readonly string _configPath;
        private readonly string _subiektConfigPath;
        private readonly string _glsConfigPath;
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
            _historyPath = Path.Combine(gryzakPath, "order_history.json");
        }

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
                        // Napraw obcięty endpoint jeśli istnieje
                        if (config.OrderDetailsEndpoint != null && config.OrderDetailsEndpoint.Contains("{order_{d}"))
                        {
                            config.OrderDetailsEndpoint = "/index.php?route=extension/module/orders/details&token=strefalicencji&order_id={order_id}&format=json";
                            SaveConfig(config); // Zapisz naprawioną konfigurację
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
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
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
            var defaultConfig = GetDefaultConfig();
            SaveConfig(defaultConfig);
        }

        private ApiConfig GetDefaultConfig()
        {
            return new ApiConfig
            {
                ApiUrl = "https://mikran.pl",
                ApiToken = "",
                ApiTimeout = 30,
                OrderListEndpoint = "/index.php?route=extension/module/orders/list&format=json",
                OrderDetailsEndpoint = "/index.php?route=extension/module/orders/details&token=strefalicencji&order_id={order_id}&format=json"
            };
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
                        if (string.IsNullOrWhiteSpace(config.DiscountCalculationMode))
                        {
                            config.DiscountCalculationMode = "percent";
                        }
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
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
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
            return new SubiektConfig
            {
                ServerAddress = "192.168.0.140",
                DatabaseName = "",
                ServerUsername = "mikran_com",
                ServerPassword = "mikran_comqwer4321",
                User = "Szef",
                Password = "zdanoszef123",
                AutoReleaseLicenseTimeoutMinutes = 0,
                DiscountCalculationMode = "percent"
            };
        }

        public GlsConfig LoadGlsConfig()
        {
            try
            {
                if (File.Exists(_glsConfigPath))
                {
                    var json = File.ReadAllText(_glsConfigPath);
                    var config = JsonSerializer.Deserialize<GlsConfig>(json);
                    if (config != null)
                    {
                        if (string.IsNullOrWhiteSpace(config.TestApiUrl))
                        {
                            config.TestApiUrl = GlsConfig.DefaultTestApiUrl;
                        }
                        if (string.IsNullOrWhiteSpace(config.ProductionApiUrl))
                        {
                            config.ProductionApiUrl = GlsConfig.DefaultProductionApiUrl;
                        }
                        if (config.TimeoutSeconds < 5 || config.TimeoutSeconds > 300)
                        {
                            config.TimeoutSeconds = 30;
                        }
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

        public void SaveGlsConfig(GlsConfig config)
        {
            try
            {
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
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
            return new GlsConfig();
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
                        // Zwróć maksymalnie 10 pozycji
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
                    OrderIds = history.Take(10).ToList() // Ogranicz do 10
                };
                var json = JsonSerializer.Serialize(historyData, new JsonSerializerOptions { WriteIndented = true });
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
                
                // Usuń duplikat jeśli istnieje
                history.RemoveAll(id => id == orderId);
                
                // Dodaj na początek
                history.Insert(0, orderId);
                
                // Ogranicz do 10 pozycji
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
    }
}

