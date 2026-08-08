using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;
using Gryzak.Views;
using static Gryzak.Services.Logger;

namespace Gryzak.Services
{
    /// <summary>
    /// Klient Subiekt REST API (zastępuje bezpośrednie zapytania MSSQL).
    /// </summary>
    public class SubiektApiService
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        private readonly ConfigService _configService;

        public SubiektApiService(ConfigService? configService = null)
        {
            _configService = configService ?? new ConfigService();
        }

        public bool IsConfigured(SubiektConfig? config = null)
        {
            config ??= _configService.LoadSubiektConfig();
            return !string.IsNullOrWhiteSpace(config.ApiBaseUrl);
        }

        public async Task<(bool Ok, string Message)> TestConnectionAsync(SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            if (!IsConfigured(config))
            {
                throw new InvalidOperationException("URL API Subiekt jest wymagany.");
            }

            using var client = CreateClient(config);
            using var response = await client.GetAsync("health", cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"API Error: {(int)response.StatusCode} {response.ReasonPhrase}. {ExtractError(body)}");
            }

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<HealthDto>>(body, JsonOptions);
            var status = envelope?.Data?.Status ?? "?";
            var sql = envelope?.Data?.SqlConfigured == true ? "SQL OK" : "SQL nie skonfigurowane";
            return (true, $"status={status}, {sql}");
        }

        public async Task<List<VatRate>> GetVatRatesAsync(SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            using var response = await client.GetAsync("vat", cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<VatDto>>>(body, JsonOptions);
            var list = new List<VatRate>();
            if (envelope?.Data == null)
            {
                return list;
            }

            foreach (var item in envelope.Data)
            {
                if (string.IsNullOrWhiteSpace(item.Vat_Symbol))
                {
                    continue;
                }

                list.Add(new VatRate
                {
                    VatId = item.Vat_Id,
                    VatSymbol = item.Vat_Symbol,
                    VatStawka = item.Vat_Stawka
                });
            }

            return list;
        }

        public async Task<int?> FindCountryIdByNameAsync(string countryName, SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(countryName))
            {
                return null;
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            var url = $"kraje?nazwa={Uri.EscapeDataString(countryName.Trim())}";
            using var response = await client.GetAsync(url, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<CountryDto>>>(body, JsonOptions);
            var first = envelope?.Data?.FirstOrDefault();
            return first?.Pa_Id;
        }

        public async Task<List<KontrahentItem>> SearchKontrahenciAsync(
            string? email = null,
            string? customerName = null,
            string? nip = null,
            int pageSize = 20,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            var byId = new Dictionary<int, KontrahentItem>();

            async Task FetchAsync(string query)
            {
                using var client = CreateClient(config);
                using var response = await client.GetAsync($"kontrahenci?{query}", cancellationToken).ConfigureAwait(false);
                var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                EnsureSuccess(response, body);

                var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<KontrahentDto>>>(body, JsonOptions);
                if (envelope?.Data == null)
                {
                    return;
                }

                foreach (var row in envelope.Data)
                {
                    if (row.Kh_Id <= 0 || byId.ContainsKey(row.Kh_Id))
                    {
                        continue;
                    }

                    byId[row.Kh_Id] = new KontrahentItem
                    {
                        Id = row.Kh_Id,
                        Symbol = row.Kh_Symbol ?? "",
                        NazwaPelna = row.Adr_NazwaPelna ?? row.Adr_Nazwa ?? "",
                        Email = row.Kh_EMail ?? "",
                        NIP = row.Adr_NIP ?? "",
                        Adres = row.Adr_Adres ?? row.Adr_Ulica ?? "",
                        Miejscowosc = row.Adr_Miejscowosc ?? "",
                        Kod = row.Adr_Kod ?? ""
                    };
                }
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(email))
                {
                    await FetchAsync($"email={Uri.EscapeDataString(email.Trim())}&pageSize={pageSize}").ConfigureAwait(false);
                }

                var search = !string.IsNullOrWhiteSpace(customerName)
                    ? customerName.Trim()
                    : (!string.IsNullOrWhiteSpace(nip) ? nip.Trim() : null);

                if (!string.IsNullOrWhiteSpace(search))
                {
                    await FetchAsync($"search={Uri.EscapeDataString(search)}&pageSize={pageSize}").ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                Error(ex, "SubiektApiService", "Błąd podczas wyszukiwania kontrahentów");
            }

            return byId.Values.Take(pageSize).ToList();
        }

        public async Task<int?> GetZkIdByNumerOryginalnyAsync(string numerOryginalny, SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(numerOryginalny))
            {
                return null;
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            var url = $"documents/zk?numerOryginalny={Uri.EscapeDataString(numerOryginalny.Trim())}&pageSize=1";
            using var response = await client.GetAsync(url, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<ZkDocumentDto>>>(body, JsonOptions);
            var first = envelope?.Data?.FirstOrDefault();
            return first?.Dok_Id;
        }

        public async Task<Dictionary<string, string>> GetZkByNumerOryginalnyInAsync(
            IEnumerable<string> numeryOryginalne,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            var wynik = new Dictionary<string, string>();
            var numery = numeryOryginalne
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim())
                .Distinct()
                .ToList();

            if (numery.Count == 0)
            {
                return wynik;
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            const int chunkSize = 200;
            for (int offset = 0; offset < numery.Count; offset += chunkSize)
            {
                var chunk = numery.Skip(offset).Take(chunkSize).ToList();
                var csv = string.Join(",", chunk);
                using var client = CreateClient(config);
                var url = $"documents/zk?numerOryginalnyIn={Uri.EscapeDataString(csv)}";
                using var response = await client.GetAsync(url, cancellationToken).ConfigureAwait(false);
                var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                EnsureSuccess(response, body);

                var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<ZkOrygLookupDto>>>(body, JsonOptions);
                if (envelope?.Data == null)
                {
                    continue;
                }

                foreach (var row in envelope.Data)
                {
                    if (string.IsNullOrWhiteSpace(row.Dok_NrPelnyOryg))
                    {
                        continue;
                    }

                    var key = row.Dok_NrPelnyOryg.Trim();
                    if (!wynik.ContainsKey(key))
                    {
                        wynik[key] = row.Dok_NrPelny?.Trim() ?? "";
                    }
                }
            }

            return wynik;
        }

        public async Task<List<SubiektUserDto>> GetUsersAsync(SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            var users = new List<SubiektUserDto>();
            int page = 1;
            int totalPages = 1;

            using var client = CreateClient(config);
            while (page <= totalPages)
            {
                using var response = await client.GetAsync($"uzytkownicy?page={page}&pageSize=100", cancellationToken).ConfigureAwait(false);
                var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                EnsureSuccess(response, body);

                var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<SubiektUserDto>>>(body, JsonOptions);
                if (envelope?.Data != null)
                {
                    users.AddRange(envelope.Data);
                }

                totalPages = Math.Max(1, envelope?.Pagination?.TotalPages ?? 1);
                page++;
            }

            return users
                .OrderBy(u => u.Uz_Nazwisko ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(u => u.Uz_Imie ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public async Task<List<(int Id, string Nazwa)>> GetPluMismatchProductsAsync(SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            var result = new List<(int Id, string Nazwa)>();
            int page = 1;
            int totalPages = 1;

            using var client = CreateClient(config);
            while (page <= totalPages)
            {
                using var response = await client.GetAsync($"products?pluMismatch=true&page={page}&pageSize=200", cancellationToken).ConfigureAwait(false);
                var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                EnsureSuccess(response, body);

                var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<ProductSlimDto>>>(body, JsonOptions);
                if (envelope?.Data != null)
                {
                    foreach (var p in envelope.Data)
                    {
                        result.Add((p.Tw_Id, p.Tw_Nazwa ?? ""));
                    }
                }

                totalPages = Math.Max(1, envelope?.Pagination?.TotalPages ?? 1);
                page++;
            }

            return result;
        }

        public async Task<bool> FixPluAsync(int twId, SubiektConfig? config = null, CancellationToken cancellationToken = default)
        {
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            using var content = new StringContent(
                JsonSerializer.Serialize(new { ids = new[] { twId } }),
                Encoding.UTF8,
                "application/json");

            using var response = await client.PostAsync("products/fix-plu", content, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<FixPluResultDto>>(body, JsonOptions);
            return (envelope?.Data?.UpdatedCount ?? 0) > 0
                   || (envelope?.Data?.UpdatedIds?.Contains(twId) ?? false);
        }

        public T RunSync<T>(Func<Task<T>> action)
        {
            return action().ConfigureAwait(false).GetAwaiter().GetResult();
        }

        private static HttpClient CreateClient(SubiektConfig config)
        {
            var baseUrl = config.ApiBaseUrl.Trim().TrimEnd('/') + "/";
            var client = new HttpClient
            {
                BaseAddress = new Uri(baseUrl, UriKind.Absolute),
                Timeout = TimeSpan.FromSeconds(30)
            };
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json");

            if (!string.IsNullOrWhiteSpace(config.ApiKey))
            {
                client.DefaultRequestHeaders.Remove("X-Api-Key");
                client.DefaultRequestHeaders.Add("X-Api-Key", config.ApiKey.Trim());
            }

            return client;
        }

        private static void EnsureConfigured(SubiektConfig config)
        {
            if (string.IsNullOrWhiteSpace(config.ApiBaseUrl))
            {
                throw new InvalidOperationException("Brak URL API Subiekt w konfiguracji.");
            }
        }

        private static void EnsureSuccess(HttpResponseMessage response, string body)
        {
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            throw new Exception($"API Error: {(int)response.StatusCode} {response.ReasonPhrase}. {ExtractError(body)}");
        }

        private static string ExtractError(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return "";
            }

            try
            {
                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("error", out var err))
                {
                    return err.GetString() ?? body;
                }
            }
            catch
            {
                // ignore
            }

            return body.Length > 300 ? body.Substring(0, 300) : body;
        }

        private class ApiEnvelope<T>
        {
            public T? Data { get; set; }
            public PaginationDto? Pagination { get; set; }
        }

        private class PaginationDto
        {
            public int Page { get; set; }
            public int PageSize { get; set; }
            public int TotalCount { get; set; }
            public int TotalPages { get; set; }
        }

        private class HealthDto
        {
            public string? Status { get; set; }
            public bool SqlConfigured { get; set; }
        }

        private class VatDto
        {
            public int Vat_Id { get; set; }
            public string? Vat_Symbol { get; set; }
            public decimal Vat_Stawka { get; set; }
        }

        private class CountryDto
        {
            public int Pa_Id { get; set; }
            public string? Pa_Nazwa { get; set; }
        }

        private class KontrahentDto
        {
            public int Kh_Id { get; set; }
            public string? Kh_Symbol { get; set; }
            public string? Kh_EMail { get; set; }
            public string? Adr_Nazwa { get; set; }
            public string? Adr_NazwaPelna { get; set; }
            public string? Adr_NIP { get; set; }
            public string? Adr_Adres { get; set; }
            public string? Adr_Ulica { get; set; }
            public string? Adr_Miejscowosc { get; set; }
            public string? Adr_Kod { get; set; }
        }

        private class ZkDocumentDto
        {
            public int Dok_Id { get; set; }
            public string? Dok_NrPelnyOryg { get; set; }
        }

        private class ZkOrygLookupDto
        {
            public string? Dok_NrPelnyOryg { get; set; }
            public string? Dok_NrPelny { get; set; }
        }

        private class ProductSlimDto
        {
            public int Tw_Id { get; set; }
            public string? Tw_Nazwa { get; set; }
        }

        private class FixPluResultDto
        {
            public int UpdatedCount { get; set; }
            public List<int>? UpdatedIds { get; set; }
        }

        public class SubiektUserDto
        {
            public int Uz_Id { get; set; }
            public string? Uz_Imie { get; set; }
            public string? Uz_Nazwisko { get; set; }
        }
    }
}
