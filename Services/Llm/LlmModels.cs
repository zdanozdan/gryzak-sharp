using System;
using System.Collections.Generic;

namespace Gryzak.Services.Llm
{
    public enum LlmRole
    {
        System,
        User,
        Assistant
    }

    public sealed class LlmMessage
    {
        public LlmRole Role { get; set; }
        public string Content { get; set; } = "";

        public static LlmMessage User(string content) => new() { Role = LlmRole.User, Content = content };
        public static LlmMessage Assistant(string content) => new() { Role = LlmRole.Assistant, Content = content };
        public static LlmMessage System(string content) => new() { Role = LlmRole.System, Content = content };
    }

    /// <summary>Wspólne żądanie completion niezależne od dostawcy.</summary>
    public sealed class LlmCompletionRequest
    {
        public string? SystemPrompt { get; set; }
        public List<LlmMessage> Messages { get; set; } = new();
        public double? Temperature { get; set; }
        public int? MaxOutputTokens { get; set; }
    }

    /// <summary>Wspólna odpowiedź completion niezależna od dostawcy.</summary>
    public sealed class LlmCompletionResult
    {
        public bool Success { get; set; }
        public string Text { get; set; } = "";
        public string? ErrorMessage { get; set; }
        public string? FinishReason { get; set; }
        public string ProviderId { get; set; } = "";
        public string Model { get; set; } = "";
        public int? PromptTokens { get; set; }
        public int? CompletionTokens { get; set; }
        public TimeSpan Duration { get; set; }

        public static LlmCompletionResult Ok(string text, string providerId, string model, TimeSpan duration) =>
            new()
            {
                Success = true,
                Text = text ?? "",
                ProviderId = providerId,
                Model = model,
                Duration = duration
            };

        public static LlmCompletionResult Fail(string error, string providerId, string model, TimeSpan duration) =>
            new()
            {
                Success = false,
                ErrorMessage = error,
                ProviderId = providerId,
                Model = model,
                Duration = duration
            };
    }
}
