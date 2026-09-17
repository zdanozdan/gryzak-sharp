using System;

namespace Gryzak.Models
{
    public static class AiProviders
    {
        public const string GoogleGemini = "google_gemini";

        public static string GetDisplayName(string? provider) =>
            string.Equals(provider, GoogleGemini, StringComparison.OrdinalIgnoreCase)
                ? "Google Gemini"
                : (string.IsNullOrWhiteSpace(provider) ? "Google Gemini" : provider);
    }

    /// <summary>Ustawienia dostawcy LLM.</summary>
    public class AiConfig
    {
        public const string DefaultGeminiModel = "gemini-3.5-flash-lite";
        public const string DefaultGeminiApiBaseUrl = "https://generativelanguage.googleapis.com";

        /// <summary>Identyfikator dostawcy, np. google_gemini.</summary>
        public string Provider { get; set; } = AiProviders.GoogleGemini;

        /// <summary>Nazwa modelu u dostawcy (np. gemini-3.5-flash-lite).</summary>
        public string Model { get; set; } = DefaultGeminiModel;

        /// <summary>Klucz API dostawcy.</summary>
        public string ApiKey { get; set; } = "";

        /// <summary>
        /// Bazowy URL API. Pusty = domyślny endpoint dostawcy
        /// (dla Gemini: generativelanguage.googleapis.com).
        /// </summary>
        public string ApiBaseUrl { get; set; } = "";

        /// <summary>Temperatura generowania (0–2).</summary>
        public double Temperature { get; set; } = 0.7;

        /// <summary>Maksymalna liczba tokenów w odpowiedzi (0 = domyślna dostawcy).</summary>
        public int MaxOutputTokens { get; set; } = 8192;

        /// <summary>Timeout HTTP na wywołanie LLM (sekundy).</summary>
        public int TimeoutSeconds { get; set; } = 60;

        public void Normalize()
        {
            if (string.IsNullOrWhiteSpace(Provider))
            {
                Provider = AiProviders.GoogleGemini;
            }

            Provider = Provider.Trim().ToLowerInvariant();
            Model = string.IsNullOrWhiteSpace(Model) ? DefaultGeminiModel : Model.Trim();
            ApiKey ??= "";
            ApiBaseUrl = (ApiBaseUrl ?? "").Trim().TrimEnd('/');

            if (double.IsNaN(Temperature) || Temperature < 0)
            {
                Temperature = 0;
            }
            else if (Temperature > 2)
            {
                Temperature = 2;
            }

            if (MaxOutputTokens < 0)
            {
                MaxOutputTokens = 0;
            }

            TimeoutSeconds = ClampTimeout(TimeoutSeconds);
        }

        public string GetEffectiveApiBaseUrl()
        {
            if (!string.IsNullOrWhiteSpace(ApiBaseUrl))
            {
                return ApiBaseUrl.Trim().TrimEnd('/');
            }

            return string.Equals(Provider, AiProviders.GoogleGemini, StringComparison.OrdinalIgnoreCase)
                ? DefaultGeminiApiBaseUrl
                : "";
        }

        private static int ClampTimeout(int seconds)
        {
            if (seconds < 5)
            {
                return 5;
            }

            if (seconds > 300)
            {
                return 300;
            }

            return seconds;
        }
    }
}
