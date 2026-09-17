using System;
using Gryzak.Models;

namespace Gryzak.Services.Llm
{
    /// <summary>Tworzy klienta LLM na podstawie AiConfig (provider + model + klucz).</summary>
    public static class LlmClientFactory
    {
        public static ILlmClient Create(AiConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            config.Normalize();

            return config.Provider switch
            {
                AiProviders.GoogleGemini => new GeminiLlmClient(config),
                // AiProviders.Anthropic => new ClaudeLlmClient(config),
                // AiProviders.Xai => new GrokLlmClient(config),
                _ => throw new NotSupportedException(
                    $"Nieobsługiwany dostawca AI: '{config.Provider}'.")
            };
        }
    }
}
