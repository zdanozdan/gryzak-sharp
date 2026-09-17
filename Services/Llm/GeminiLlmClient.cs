using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;

namespace Gryzak.Services.Llm
{
    /// <summary>Klient Google Gemini (generateContent).</summary>
    public sealed class GeminiLlmClient : ILlmClient
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        private readonly AiConfig _config;
        private readonly HttpClient _httpClient;

        public string ProviderId => AiProviders.GoogleGemini;
        public string Model => _config.Model;

        public GeminiLlmClient(AiConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _config.Normalize();

            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(_config.TimeoutSeconds)
            };
            _httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<LlmCompletionResult> CompleteAsync(
            LlmCompletionRequest request,
            CancellationToken cancellationToken = default)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var sw = Stopwatch.StartNew();

            if (string.IsNullOrWhiteSpace(_config.ApiKey))
            {
                return LlmCompletionResult.Fail(
                    "Brak klucza API Gemini.",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }

            if (string.IsNullOrWhiteSpace(_config.Model))
            {
                return LlmCompletionResult.Fail(
                    "Brak nazwy modelu.",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }

            var baseUrl = _config.GetEffectiveApiBaseUrl();
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                return LlmCompletionResult.Fail(
                    "Brak bazowego URL API Gemini.",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }

            try
            {
                var url =
                    $"{baseUrl}/v1beta/models/{Uri.EscapeDataString(_config.Model)}:generateContent";

                using var httpRequest = new HttpRequestMessage(HttpMethod.Post, url);
                httpRequest.Headers.TryAddWithoutValidation("x-goog-api-key", _config.ApiKey);
                httpRequest.Content = new StringContent(
                    BuildRequestBody(request),
                    Encoding.UTF8,
                    "application/json");

                using var response = await _httpClient
                    .SendAsync(httpRequest, cancellationToken)
                    .ConfigureAwait(false);
                var body = await response.Content.ReadAsStringAsync(cancellationToken)
                    .ConfigureAwait(false);
                sw.Stop();

                if (!response.IsSuccessStatusCode)
                {
                    return LlmCompletionResult.Fail(
                        FormatHttpError((int)response.StatusCode, body),
                        ProviderId,
                        Model,
                        sw.Elapsed);
                }

                return ParseSuccess(body, sw.Elapsed);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                sw.Stop();
                return LlmCompletionResult.Fail(
                    "Timeout — odpowiedź Gemini przekroczyła limit czasu.",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }
            catch (OperationCanceledException)
            {
                sw.Stop();
                return LlmCompletionResult.Fail(
                    "Anulowano.",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }
            catch (HttpRequestException ex)
            {
                sw.Stop();
                return LlmCompletionResult.Fail(
                    $"Błąd sieci: {ex.Message}",
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }
            catch (Exception ex)
            {
                sw.Stop();
                return LlmCompletionResult.Fail(
                    ex.Message,
                    ProviderId,
                    Model,
                    sw.Elapsed);
            }
        }

        private string BuildRequestBody(LlmCompletionRequest request)
        {
            var root = new JsonObject();

            var systemText = request.SystemPrompt;
            if (string.IsNullOrWhiteSpace(systemText))
            {
                systemText = request.Messages
                    .FirstOrDefault(m => m.Role == LlmRole.System)
                    ?.Content;
            }

            if (!string.IsNullOrWhiteSpace(systemText))
            {
                root["systemInstruction"] = new JsonObject
                {
                    ["parts"] = new JsonArray
                    {
                        new JsonObject { ["text"] = systemText }
                    }
                };
            }

            var contents = new JsonArray();
            foreach (var message in request.Messages)
            {
                if (message.Role == LlmRole.System)
                {
                    continue;
                }

                var role = message.Role == LlmRole.Assistant ? "model" : "user";
                contents.Add(new JsonObject
                {
                    ["role"] = role,
                    ["parts"] = new JsonArray
                    {
                        new JsonObject { ["text"] = message.Content ?? "" }
                    }
                });
            }

            if (contents.Count == 0)
            {
                contents.Add(new JsonObject
                {
                    ["role"] = "user",
                    ["parts"] = new JsonArray
                    {
                        new JsonObject { ["text"] = "" }
                    }
                });
            }

            root["contents"] = contents;

            var temperature = request.Temperature ?? _config.Temperature;
            var maxTokens = request.MaxOutputTokens ?? _config.MaxOutputTokens;

            var generationConfig = new JsonObject
            {
                ["temperature"] = temperature
            };
            if (maxTokens > 0)
            {
                generationConfig["maxOutputTokens"] = maxTokens;
            }

            root["generationConfig"] = generationConfig;

            return root.ToJsonString(JsonOptions);
        }

        private LlmCompletionResult ParseSuccess(string body, TimeSpan duration)
        {
            try
            {
                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                if (root.TryGetProperty("error", out var errorEl))
                {
                    var msg = errorEl.TryGetProperty("message", out var m)
                        ? m.GetString()
                        : errorEl.ToString();
                    return LlmCompletionResult.Fail(
                        msg ?? "Nieznany błąd Gemini.",
                        ProviderId,
                        Model,
                        duration);
                }

                var text = ExtractText(root);
                var finishReason = ExtractFinishReason(root);
                var result = LlmCompletionResult.Ok(text, ProviderId, Model, duration);
                result.FinishReason = finishReason;

                if (root.TryGetProperty("usageMetadata", out var usage))
                {
                    if (usage.TryGetProperty("promptTokenCount", out var prompt)
                        && prompt.TryGetInt32(out var promptTokens))
                    {
                        result.PromptTokens = promptTokens;
                    }

                    if (usage.TryGetProperty("candidatesTokenCount", out var completion)
                        && completion.TryGetInt32(out var completionTokens))
                    {
                        result.CompletionTokens = completionTokens;
                    }
                }

                if (string.IsNullOrWhiteSpace(result.Text))
                {
                    return LlmCompletionResult.Fail(
                        string.IsNullOrWhiteSpace(finishReason)
                            ? "Pusta odpowiedź modelu."
                            : $"Pusta odpowiedź modelu (finishReason: {finishReason}).",
                        ProviderId,
                        Model,
                        duration);
                }

                return result;
            }
            catch (JsonException ex)
            {
                return LlmCompletionResult.Fail(
                    $"Niepoprawna odpowiedź JSON Gemini: {ex.Message}",
                    ProviderId,
                    Model,
                    duration);
            }
        }

        private static string ExtractText(JsonElement root)
        {
            if (!root.TryGetProperty("candidates", out var candidates)
                || candidates.ValueKind != JsonValueKind.Array
                || candidates.GetArrayLength() == 0)
            {
                return "";
            }

            var candidate = candidates[0];
            if (!candidate.TryGetProperty("content", out var content)
                || !content.TryGetProperty("parts", out var parts)
                || parts.ValueKind != JsonValueKind.Array)
            {
                return "";
            }

            var sb = new StringBuilder();
            foreach (var part in parts.EnumerateArray())
            {
                // Pomiń części „thinking” jeśli API je zwróci — bierzemy zwykły text.
                if (part.TryGetProperty("thought", out var thought)
                    && thought.ValueKind == JsonValueKind.True)
                {
                    continue;
                }

                if (part.TryGetProperty("text", out var textEl))
                {
                    var chunk = textEl.GetString();
                    if (!string.IsNullOrEmpty(chunk))
                    {
                        if (sb.Length > 0)
                        {
                            sb.AppendLine();
                        }

                        sb.Append(chunk);
                    }
                }
            }

            return sb.ToString().Trim();
        }

        private static string? ExtractFinishReason(JsonElement root)
        {
            if (!root.TryGetProperty("candidates", out var candidates)
                || candidates.ValueKind != JsonValueKind.Array
                || candidates.GetArrayLength() == 0)
            {
                return null;
            }

            if (candidates[0].TryGetProperty("finishReason", out var reason))
            {
                return reason.GetString();
            }

            return null;
        }

        private static string FormatHttpError(int statusCode, string body)
        {
            try
            {
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("error", out var error)
                    && error.TryGetProperty("message", out var message))
                {
                    var msg = message.GetString();
                    if (!string.IsNullOrWhiteSpace(msg))
                    {
                        return $"HTTP {statusCode}: {msg}";
                    }
                }
            }
            catch (JsonException)
            {
                // ignore — zwróć skrócone body
            }

            var snippet = string.IsNullOrWhiteSpace(body)
                ? "(brak treści)"
                : body.Length > 300 ? body.Substring(0, 300) + "…" : body;
            return $"HTTP {statusCode}: {snippet}";
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }
    }
}
