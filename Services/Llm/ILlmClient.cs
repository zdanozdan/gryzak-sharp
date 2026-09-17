using System;
using System.Threading;
using System.Threading.Tasks;

namespace Gryzak.Services.Llm
{
    /// <summary>
    /// Wspólny kontrakt klienta LLM. Nowe dostawcy (Claude, Grok, …)
    /// implementują ten interfejs — UI i logika biznesowa wołają tylko CompleteAsync.
    /// </summary>
    public interface ILlmClient : IDisposable
    {
        string ProviderId { get; }
        string Model { get; }

        Task<LlmCompletionResult> CompleteAsync(
            LlmCompletionRequest request,
            CancellationToken cancellationToken = default);
    }
}
