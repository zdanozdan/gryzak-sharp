using System;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;

namespace Gryzak.Services.Llm
{
    public sealed class GlsAddressExtractionResult
    {
        public bool Success { get; init; }
        public DocumentAiAddress? Address { get; init; }
        public string? ErrorMessage { get; init; }
        public string RawResponse { get; init; } = "";
        public bool Found { get; init; }

        public static GlsAddressExtractionResult Ok(DocumentAiAddress address, string raw) =>
            new()
            {
                Success = true,
                Found = true,
                Address = address,
                RawResponse = raw
            };

        public static GlsAddressExtractionResult NotFound(string? reason, string raw) =>
            new()
            {
                Success = true,
                Found = false,
                ErrorMessage = string.IsNullOrWhiteSpace(reason)
                    ? "LLM nie znalazł adresu dostawy w uwagach."
                    : reason,
                RawResponse = raw
            };

        public static GlsAddressExtractionResult Fail(string error, string raw = "") =>
            new()
            {
                Success = false,
                Found = false,
                ErrorMessage = error,
                RawResponse = raw
            };
    }

    /// <summary>
    /// Wyciąga adres GLS z tekstu uwag dokumentu przez wspólne API <see cref="ILlmClient"/>.
    /// </summary>
    public sealed class GlsAddressExtractionService
    {
        private static readonly Regex JsonFenceRegex = new(
            @"```(?:json)?\s*(.*?)```",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private const string SystemPrompt =
            """
            Jesteś asystentem logistyki firmy kurierskiej GLS w Polsce.
            Z tekstu uwag dokumentu handlowego (Subiekt GT) wyodrębniasz WYŁĄCZNIE adres DOSTAWY
            (odbiorca paczki): imię/nazwa, ulica z numerem, kod pocztowy, miasto, kraj, telefon/email jeśli przy adresie.

            Adresy mogą pochodzić z różnych krajów UE (nie tylko Polska) — np. DE, CZ, SK, AT, LT, LV, EE,
            NL, BE, FR, IT, ES, HU, RO, BG, HR, SI, IE, DK, SE, FI, PT, GR, LU, MT, CY oraz spoza UE jeśli widać w tekście.
            Ustal kraj z nazwy kraju, kodu ISO, formatu kodu pocztowego lub języka/formy adresu.
            Nie zakładaj automatycznie Polski, jeśli tekst wskazuje inny kraj.

            W uwagach często jest dużo zbędnych danych zamówienia — IGNORUJ je całkowicie.
            Nie wstawiaj ich do żadnego pola JSON (ani name*, ani street, ani notes).

            Ignoruj m.in.:
            - rozmiary, wymiary, wagę, ilości, SKU, kody produktów, warianty
            - nazwy produktów, opisy, uwagi techniczne, komentarze do towaru
            - ceny, rabaty, formy płatności, statusy zamówienia
            - numery faktur / WZ / ZK / PA (chyba że to jedyny sensowny notes — wtedy tylko sam numer)
            - NIP, REGON, KRS, VAT ID, dane księgowe
            - adres siedziby / rozliczeniowy, jeśli w tekście jest osobny adres dostawy / wysyłki
            - inne „dziwne” wpisy niezwiązane z lokalizacją odbiorcy paczki

            Zwracasz WYŁĄCZNIE jeden obiekt JSON (bez markdown, bez komentarzy) o polach:
            {
              "found": true|false,
              "confidence": "high"|"medium"|"low",
              "reason": "krótko po polsku gdy found=false lub niska pewność",
              "name1": "nazwa odbiorcy / firma — max 40 znaków, wymagane gdy found=true",
              "name2": "opcjonalnie druga linia nazwy — max 40",
              "name3": "opcjonalnie trzecia linia — max 40",
              "street": "ulica i numer domu/lokalu — max 40, wymagane gdy found=true",
              "zipCode": "kod pocztowy w formacie właściwym dla kraju (PL: XX-XXX; DE: 5 cyfr; itd.)",
              "city": "miasto — max 40",
              "country": "kod ISO2 kraju (np. PL, DE, CZ, SK); PL tylko gdy naprawdę Polska",
              "phone": "telefon kontaktowy odbiorcy jeśli jest przy adresie (z kierunkowym jeśli w tekście)",
              "contact": "email odbiorcy jeśli jest przy adresie",
              "notes": "puste, chyba że jest krótki nr zamówienia sklepu — max 40; NIGDY rozmiary/produkty"
            }

            Zasady:
            - found=false gdy brak wiarygodnego adresu dostawy (ulica + kod + miasto + odbiorca).
            - Preferuj adres dostawy / wysyłki nad siedzibą firmy.
            - Zachowaj lokalny format kodu pocztowego kraju — nie przerabiaj zagranicznych kodów na polski XX-XXX.
            - Dla Polski normalizuj kod do XX-XXX.
            - Nie wymyślaj ulicy, kodu, miasta ani kraju — tylko to, co wynika z tekstu.
            - Skracaj pola do limitów GLS (40 znaków dla name/street/city).
            - Jeśli w tekście jest adres i obok śmieci zamówieniowe — weź tylko adres.
            """;

        private const string NormalizeShopSystemPrompt =
            """
            Jesteś asystentem logistyki GLS. Dostajesz surowy adres wysyłki ze sklepu internetowego
            (często nieuporządkowany: złe limity pól, firma+osoba w jednej linii, kod bez myślnika,
            mieszanka PL/UE). Znormalizuj go do pól etykiety GLS.

            Adresy mogą być z Polski lub innych krajów UE — ustal country (ISO2) i format kodu pocztowego
            właściwy dla kraju. Nie zakładaj PL, jeśli dane wskazują inny kraj.

            Zwracasz WYŁĄCZNIE jeden obiekt JSON (bez markdown) o polach:
            {
              "found": true|false,
              "confidence": "high"|"medium"|"low",
              "reason": "krótko po polsku gdy found=false",
              "name1": "firma lub nazwisko — max 40",
              "name2": "osoba kontaktowa / druga linia — max 40",
              "name3": "opcjonalnie — max 40",
              "street": "ulica i numer — max 40",
              "zipCode": "kod pocztowy (PL: XX-XXX)",
              "city": "miasto — max 40",
              "country": "ISO2",
              "phone": "telefon — max 25",
              "contact": "email — max 40",
              "notes": "opcjonalnie nr zamówienia — max 40"
            }

            Zasady:
            - Zachowaj sens danych — nie wymyślaj ulicy/miasta/kodu.
            - Popraw oczywiste formatowanie (kod PL, zbędne spacje, kolejność ulica/numer).
            - Jeśli jest firma i osoba: name1=firma, name2=osoba (oba ≤40 znaków).
            - found=false tylko gdy po normalizacji brakuje krytycznych pól (name1, street, zip, city).
            - Ignoruj śmieci niezwiązane z adresem.
            """;

        private readonly ILlmClient _client;

        public GlsAddressExtractionService(ILlmClient client)
        {
            _client = client ?? throw new ArgumentNullException(nameof(client));
        }

        public static string ComputeUwagiHash(string? uwagi)
        {
            var normalized = NormalizeUwagi(uwagi);
            if (normalized.Length == 0)
            {
                return "";
            }

            var hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
            return Convert.ToHexString(hash).ToLowerInvariant();
        }

        public async Task<GlsAddressExtractionResult> ExtractAsync(
            string? uwagi,
            CancellationToken cancellationToken = default)
        {
            var text = NormalizeUwagi(uwagi);
            if (text.Length == 0)
            {
                return GlsAddressExtractionResult.Fail("Brak tekstu uwag do analizy.");
            }

            return await CompleteAddressAsync(
                    SystemPrompt,
                    "Wyodrębnij adres dostawy GLS z poniższych uwag dokumentu.\n\n---\n"
                    + text
                    + "\n---",
                    text,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        /// <summary>Normalizuje adres wysyłki ze sklepu do limitów i formatu GLS.</summary>
        public async Task<GlsAddressExtractionResult> NormalizeShopAddressAsync(
            string shopAddressText,
            CancellationToken cancellationToken = default)
        {
            var text = NormalizeUwagi(shopAddressText);
            if (text.Length == 0)
            {
                return GlsAddressExtractionResult.Fail("Brak danych adresu sklepu do normalizacji.");
            }

            return await CompleteAddressAsync(
                    NormalizeShopSystemPrompt,
                    "Znormalizuj poniższy adres wysyłki ze sklepu do pól GLS.\n\n---\n"
                    + text
                    + "\n---",
                    text,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        private async Task<GlsAddressExtractionResult> CompleteAddressAsync(
            string systemPrompt,
            string userMessage,
            string sourceForHash,
            CancellationToken cancellationToken)
        {
            var request = new LlmCompletionRequest
            {
                SystemPrompt = systemPrompt,
                Messages = { LlmMessage.User(userMessage) },
                Temperature = 0,
                MaxOutputTokens = 1024
            };

            var completion = await _client.CompleteAsync(request, cancellationToken)
                .ConfigureAwait(false);
            if (!completion.Success)
            {
                return GlsAddressExtractionResult.Fail(
                    completion.ErrorMessage ?? "Błąd wywołania LLM.",
                    completion.Text);
            }

            return ParseResponse(completion.Text, sourceForHash);
        }

        internal static GlsAddressExtractionResult ParseResponse(string rawText, string uwagiSource)
        {
            if (!TryExtractJsonObject(rawText, out var json, out var parseError))
            {
                return GlsAddressExtractionResult.Fail(
                    parseError ?? "LLM nie zwrócił poprawnego JSON.",
                    rawText);
            }

            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var found = true;
                if (root.TryGetProperty("found", out var foundEl))
                {
                    found = foundEl.ValueKind switch
                    {
                        JsonValueKind.True => true,
                        JsonValueKind.False => false,
                        JsonValueKind.String => !string.Equals(
                            foundEl.GetString(), "false", StringComparison.OrdinalIgnoreCase),
                        _ => true
                    };
                }

                var reason = ReadString(root, "reason");
                if (!found)
                {
                    return GlsAddressExtractionResult.NotFound(reason, rawText);
                }

                var country = NormalizeCountry(ReadString(root, "country"));
                var address = new DocumentAiAddress
                {
                    Name1 = Truncate(ReadString(root, "name1"), 40),
                    Name2 = Truncate(ReadString(root, "name2"), 40),
                    Name3 = Truncate(ReadString(root, "name3"), 40),
                    Street = Truncate(ReadString(root, "street"), 40),
                    ZipCode = NormalizeZip(
                        ReadString(root, "zipCode", "zip_code", "postalCode"),
                        country),
                    City = Truncate(ReadString(root, "city"), 40),
                    Country = country,
                    Phone = Truncate(ReadString(root, "phone"), 25),
                    Contact = Truncate(ReadString(root, "contact", "email"), 40),
                    Notes = Truncate(ReadString(root, "notes"), 40),
                    Confidence = ReadString(root, "confidence"),
                    RawJson = json,
                    UwagiHash = ComputeUwagiHash(uwagiSource)
                };

                if (!address.IsComplete)
                {
                    return GlsAddressExtractionResult.NotFound(
                        string.IsNullOrWhiteSpace(reason)
                            ? "LLM zwrócił niekompletny adres (wymagane: name1, street, zipCode, city)."
                            : reason,
                        rawText);
                }

                return GlsAddressExtractionResult.Ok(address, rawText);
            }
            catch (JsonException ex)
            {
                return GlsAddressExtractionResult.Fail($"Błąd JSON: {ex.Message}", rawText);
            }
        }

        private static bool TryExtractJsonObject(string raw, out string json, out string? error)
        {
            json = "";
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "Pusta odpowiedź LLM.";
                return false;
            }

            var text = raw.Trim();
            var fence = JsonFenceRegex.Match(text);
            if (fence.Success)
            {
                text = fence.Groups[1].Value.Trim();
            }

            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start < 0 || end <= start)
            {
                error = "Brak obiektu JSON w odpowiedzi LLM.";
                return false;
            }

            json = text.Substring(start, end - start + 1);
            return true;
        }

        private static string ReadString(JsonElement root, params string[] names)
        {
            foreach (var name in names)
            {
                if (root.TryGetProperty(name, out var el))
                {
                    return el.ValueKind == JsonValueKind.String
                        ? el.GetString()?.Trim() ?? ""
                        : el.ToString().Trim();
                }
            }

            return "";
        }

        private static string NormalizeUwagi(string? uwagi) =>
            (uwagi ?? "").Replace("\r\n", "\n").Trim();

        private static string Truncate(string value, int max)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= max)
            {
                return value ?? "";
            }

            return value.Substring(0, max).Trim();
        }

        private static string NormalizeZip(string zip, string country)
        {
            var trimmed = (zip ?? "").Trim();
            if (string.IsNullOrEmpty(trimmed))
            {
                return "";
            }

            // Format XX-XXX tylko dla Polski — nie psuj np. niemieckich 5-cyfrowych kodów.
            if (string.Equals(country, "PL", StringComparison.OrdinalIgnoreCase))
            {
                var digits = Regex.Replace(trimmed, @"\D", "");
                if (digits.Length == 5)
                {
                    return $"{digits.Substring(0, 2)}-{digits.Substring(2)}";
                }
            }

            return Truncate(trimmed, 10);
        }

        private static string NormalizeCountry(string country)
        {
            if (string.IsNullOrWhiteSpace(country))
            {
                return "PL";
            }

            var c = country.Trim().ToUpperInvariant();
            return c switch
            {
                "POLSKA" or "POLAND" or "POL" => "PL",
                "NIEMCY" or "GERMANY" or "DEUTSCHLAND" or "DEU" => "DE",
                "CZECHY" or "CZECH REPUBLIC" or "CZECHIA" or "ČESKO" or "CESKO" or "CZE" => "CZ",
                "SŁOWACJA" or "SLOWACJA" or "SLOVAKIA" or "SLOVENSKO" or "SVK" => "SK",
                "AUSTRIA" or "ÖSTERREICH" or "OSTERREICH" or "AUT" => "AT",
                "LITWA" or "LITHUANIA" or "LIETUVA" or "LTU" => "LT",
                "ŁOTWA" or "LOTWA" or "LATVIA" or "LATVIJA" or "LVA" => "LV",
                "ESTONIA" or "EESTI" or "EST" => "EE",
                "HOLANDIA" or "NIDERLANDY" or "NETHERLANDS" or "NEDERLAND" or "NLD" => "NL",
                "BELGIA" or "BELGIUM" or "BELGIË" or "BELGIE" or "BEL" => "BE",
                "FRANCJA" or "FRANCE" or "FRA" => "FR",
                "WŁOCHY" or "WLOCHY" or "ITALY" or "ITALIA" or "ITA" => "IT",
                "HISZPANIA" or "SPAIN" or "ESPAÑA" or "ESPANA" or "ESP" => "ES",
                "WĘGRY" or "WEGRY" or "HUNGARY" or "MAGYARORSZÁG" or "MAGYARORSZAG" or "HUN" => "HU",
                "RUMUNIA" or "ROMANIA" or "ROMÂNIA" or "ROU" => "RO",
                "BUŁGARIA" or "BULGARIA" or "БЪЛГАРИЯ" or "BGR" => "BG",
                "CHORWACJA" or "CROATIA" or "HRVATSKA" or "HRV" => "HR",
                "SŁOWENIA" or "SLOWENIA" or "SLOVENIA" or "SLOVENIJA" or "SVN" => "SI",
                "IRLANDIA" or "IRELAND" or "ÉIRE" or "EIRE" or "IRL" => "IE",
                "DANIA" or "DENMARK" or "DANMARK" or "DNK" => "DK",
                "SZWECJA" or "SWEDEN" or "SVERIGE" or "SWE" => "SE",
                "FINLANDIA" or "FINLAND" or "SUOMI" or "FIN" => "FI",
                "PORTUGALIA" or "PORTUGAL" or "PRT" => "PT",
                "GRECJA" or "GREECE" or "ΕΛΛΆΔΑ" or "HELLAS" or "GRC" => "GR",
                "LUKSEMBURG" or "LUXEMBOURG" or "LUX" => "LU",
                "MALTA" or "MLT" => "MT",
                "CYPR" or "CYPRUS" or "ΚΎΠΡΟΣ" or "CYP" => "CY",
                _ => c.Length >= 2 ? c.Substring(0, 2) : "PL"
            };
        }
    }
}
