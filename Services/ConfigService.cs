using System;
using System.IO;
using System.Text.Json;
using Gryzak.Models;

namespace Gryzak.Services
{
    public class ConfigService
    {
        private readonly string _configPath;

        public ConfigService()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var gryzakPath = Path.Combine(appDataPath, "Gryzak");
            
            if (!Directory.Exists(gryzakPath))
            {
                Directory.CreateDirectory(gryzakPath);
            }
            
            _configPath = Path.Combine(gryzakPath, "config.json");
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
                            config.OrderDetailsEndpoint = "/index.php?route=extension/module/orders&token=strefalicencji&order_id={order_id}&format=json";
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
                ApiUrl = "",
                ApiToken = "",
                ApiTimeout = 30,
                OrderListEndpoint = "/orders",
                OrderDetailsEndpoint = "/index.php?route=extension/module/orders&token=strefalicencji&order_id={order_id}&format=json"
            };
        }
    }
}

