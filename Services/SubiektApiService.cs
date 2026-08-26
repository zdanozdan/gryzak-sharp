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

        public async Task<SubiektDocumentPage> GetDocumentsAsync(
            string documentType,
            int page = 1,
            int pageSize = 20,
            string? search = null,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            var typ = NormalizeDocumentType(documentType);
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            var normalizedSearch = NormalizeDocumentSearchQuery(search);
            var hasSearch = !string.IsNullOrWhiteSpace(normalizedSearch);
            // Przy search API preferuje `limit` (domyślnie 50); pageSize=20 bywa bardzo wolne dla nazw/NIP.
            var effectivePageSize = hasSearch
                ? Math.Clamp(pageSize < 50 ? 50 : pageSize, 1, 200)
                : Math.Max(1, pageSize);

            var query = $"page={Math.Max(1, page)}&pageSize={effectivePageSize}";
            if (hasSearch)
            {
                query += $"&search={Uri.EscapeDataString(normalizedSearch)}";
                query += $"&limit={effectivePageSize}";
            }

            using var client = CreateClient(config);
            using var response = await client.GetAsync($"documents/{typ}?{query}", cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<List<DocumentDto>>>(body, JsonOptions);
            var result = new SubiektDocumentPage
            {
                Page = envelope?.Pagination?.Page ?? page,
                PageSize = envelope?.Pagination?.PageSize ?? effectivePageSize,
                TotalCount = envelope?.Pagination?.TotalCount ?? envelope?.Data?.Count ?? 0,
                TotalPages = envelope?.Pagination?.TotalPages ?? 1
            };

            if (envelope?.Data != null)
            {
                foreach (var row in envelope.Data)
                {
                    result.Items.Add(MapDocument(row, typ));
                }
            }

            return result;
        }

        /// <summary>
        /// Normalizuje frazę do GET /documents?search= (nr FS/WZ/ZK, nazwa, symbol, NIP, e-mail, kh_Id).
        /// </summary>
        public static string NormalizeDocumentSearchQuery(string? raw)
        {
            var text = (raw ?? "").Trim();
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            // NIP / same cyfry z myślnikami i spacjami → same cyfry (API lepiej trafia po NIP).
            var digits = new string(text.Where(char.IsDigit).ToArray());
            var hasLetters = text.Any(char.IsLetter);
            if (!hasLetters && digits.Length >= 6)
            {
                return digits;
            }

            return text;
        }

        public async Task<SubiektDocument?> GetDocumentByIdAsync(
            string documentType,
            int dokId,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            if (dokId <= 0)
            {
                return null;
            }

            var typ = NormalizeDocumentType(documentType);
            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            using var response = await client.GetAsync($"documents/{typ}/{dokId}", cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<DocumentDto>>(body, JsonOptions);
            if (envelope?.Data == null)
            {
                return null;
            }

            var doc = MapDocument(envelope.Data, typ);
            await ApplyAdresDostawyAsync(doc, config, cancellationToken).ConfigureAwait(false);
            return doc;
        }

        /// <summary>
        /// GET /kontrahenci/{kh_Id} — pełna kartoteka (adres podstawowy + opcjonalnie korespondencyjny i dostawy).
        /// </summary>
        public async Task<KontrahentDetails?> GetKontrahentDetailsAsync(
            int khId,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            var dto = await GetKontrahentByIdAsync(khId, config, cancellationToken).ConfigureAwait(false);
            return dto == null ? null : MapKontrahentDetails(dto);
        }

        /// <summary>
        /// GET /kontrahenci/{kh_Id} — pełna kartoteka z adresem dostawy (adr_Dostawa*), gdy aktywny.
        /// </summary>
        private async Task<KontrahentDto?> GetKontrahentByIdAsync(
            int khId,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            if (khId <= 0)
            {
                return null;
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            using var response = await client.GetAsync($"kontrahenci/{khId}", cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                return null;
            }

            EnsureSuccess(response, body);
            var envelope = JsonSerializer.Deserialize<ApiEnvelope<KontrahentDto>>(body, JsonOptions);
            return envelope?.Data;
        }

        private async Task ApplyAdresDostawyAsync(
            SubiektDocument doc,
            SubiektConfig config,
            CancellationToken cancellationToken)
        {
            if (doc.KontrahentId <= 0)
            {
                return;
            }

            try
            {
                var kh = await GetKontrahentByIdAsync(doc.KontrahentId, config, cancellationToken).ConfigureAwait(false);
                ApplyAdresDostawy(doc, kh);
            }
            catch (Exception ex)
            {
                Warning(
                    $"Nie udało się pobrać adresu dostawy kontrahenta {doc.KontrahentId}: {ex.Message}",
                    "SubiektApiService");
            }
        }

        private static void ApplyAdresDostawy(SubiektDocument doc, KontrahentDto? kh)
        {
            if (kh == null || !HasAdresDostawyFields(kh))
            {
                doc.HasAdresDostawy = false;
                return;
            }

            var ulica = FormatDostawaUlica(kh);
            var kod = kh.Adr_DostawaKod?.Trim() ?? "";
            var miasto = kh.Adr_DostawaMiejscowosc?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(ulica) && string.IsNullOrWhiteSpace(kod) && string.IsNullOrWhiteSpace(miasto))
            {
                doc.HasAdresDostawy = false;
                return;
            }

            var nazwa = !string.IsNullOrWhiteSpace(kh.Adr_DostawaNazwaPelna)
                ? kh.Adr_DostawaNazwaPelna.Trim()
                : (kh.Adr_DostawaNazwa?.Trim() ?? "");

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(ulica))
            {
                parts.Add(ulica);
            }

            var cityLine = $"{kod} {miasto}".Trim();
            if (!string.IsNullOrWhiteSpace(cityLine))
            {
                parts.Add(cityLine);
            }

            doc.HasAdresDostawy = true;
            doc.AdresDostawyNazwa = nazwa;
            doc.AdresDostawyUlica = ulica;
            doc.AdresDostawyKodPocztowy = kod;
            doc.AdresDostawyMiejscowosc = miasto;
            doc.AdresDostawyAdres = string.Join(", ", parts);
            doc.AdresDostawyKrajKod = FormatDostawaKrajKod(kh);
            doc.AdresDostawyTelefon = kh.Adr_DostawaTelefon?.Trim() ?? "";
        }

        private static bool HasAdresDostawyFields(KontrahentDto kh)
        {
            return kh.Adr_DostawaId is > 0
                   || !string.IsNullOrWhiteSpace(kh.Adr_DostawaAdres)
                   || !string.IsNullOrWhiteSpace(kh.Adr_DostawaUlica)
                   || !string.IsNullOrWhiteSpace(kh.Adr_DostawaKod)
                   || !string.IsNullOrWhiteSpace(kh.Adr_DostawaMiejscowosc);
        }

        private static string FormatDostawaUlica(KontrahentDto kh)
        {
            if (!string.IsNullOrWhiteSpace(kh.Adr_DostawaAdres))
            {
                return kh.Adr_DostawaAdres.Trim();
            }

            return kh.Adr_DostawaUlica?.Trim() ?? "";
        }

        private static string FormatDostawaKrajKod(KontrahentDto kh)
        {
            var name = kh.Adr_DostawaPanstwo?.Trim() ?? "";
            if (name.Equals("Polska", StringComparison.OrdinalIgnoreCase)
                || name.Equals("Poland", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(name))
            {
                return "PL";
            }

            if (name.Length == 2 && name.All(char.IsLetter))
            {
                return name.ToUpperInvariant();
            }

            return name;
        }

        private static KontrahentDetails MapKontrahentDetails(KontrahentDto kh)
        {
            return new KontrahentDetails
            {
                Id = kh.Kh_Id,
                Symbol = kh.Kh_Symbol?.Trim() ?? "",
                Email = kh.Kh_EMail?.Trim() ?? "",
                TypNazwa = FormatKhTyp(kh.Kh_Typ),
                Nazwa = kh.Adr_Nazwa?.Trim() ?? "",
                NazwaPelna = kh.Adr_NazwaPelna?.Trim() ?? "",
                Nip = kh.Adr_NIP?.Trim() ?? "",
                Adres = kh.Adr_Adres?.Trim() ?? "",
                Ulica = kh.Adr_Ulica?.Trim() ?? "",
                Kod = kh.Adr_Kod?.Trim() ?? "",
                Miejscowosc = kh.Adr_Miejscowosc?.Trim() ?? "",
                Poczta = kh.Adr_Poczta?.Trim() ?? "",
                Telefon = FirstNonEmpty(kh.Kh_TelefonKomorkowy, kh.Kh_Telefon, kh.Adr_Telefon),
                Panstwo = kh.Adr_Panstwo?.Trim() ?? kh.Pa_Nazwa?.Trim() ?? "",
                Wojewodztwo = kh.Adr_Wojewodztwo?.Trim() ?? "",

                HasAdresKorespondencyjny = HasAdresKorespondencyjnyFields(kh),
                KorespondencyjnyNazwa = kh.Adr_KorespondencyjnyNazwa?.Trim() ?? "",
                KorespondencyjnyNazwaPelna = kh.Adr_KorespondencyjnyNazwaPelna?.Trim() ?? "",
                KorespondencyjnyNip = kh.Adr_KorespondencyjnyNIP?.Trim() ?? "",
                KorespondencyjnyAdres = kh.Adr_KorespondencyjnyAdres?.Trim() ?? "",
                KorespondencyjnyUlica = kh.Adr_KorespondencyjnyUlica?.Trim() ?? "",
                KorespondencyjnyKod = kh.Adr_KorespondencyjnyKod?.Trim() ?? "",
                KorespondencyjnyMiejscowosc = kh.Adr_KorespondencyjnyMiejscowosc?.Trim() ?? "",
                KorespondencyjnyPoczta = kh.Adr_KorespondencyjnyPoczta?.Trim() ?? "",
                KorespondencyjnyTelefon = kh.Adr_KorespondencyjnyTelefon?.Trim() ?? "",
                KorespondencyjnyPanstwo = kh.Adr_KorespondencyjnyPanstwo?.Trim() ?? "",
                KorespondencyjnyWojewodztwo = kh.Adr_KorespondencyjnyWojewodztwo?.Trim() ?? "",

                HasAdresDostawy = HasAdresDostawyFields(kh),
                DostawaNazwa = kh.Adr_DostawaNazwa?.Trim() ?? "",
                DostawaNazwaPelna = kh.Adr_DostawaNazwaPelna?.Trim() ?? "",
                DostawaNip = kh.Adr_DostawaNIP?.Trim() ?? "",
                DostawaAdres = kh.Adr_DostawaAdres?.Trim() ?? "",
                DostawaUlica = kh.Adr_DostawaUlica?.Trim() ?? "",
                DostawaKod = kh.Adr_DostawaKod?.Trim() ?? "",
                DostawaMiejscowosc = kh.Adr_DostawaMiejscowosc?.Trim() ?? "",
                DostawaPoczta = kh.Adr_DostawaPoczta?.Trim() ?? "",
                DostawaTelefon = kh.Adr_DostawaTelefon?.Trim() ?? "",
                DostawaPanstwo = kh.Adr_DostawaPanstwo?.Trim() ?? "",
                DostawaWojewodztwo = kh.Adr_DostawaWojewodztwo?.Trim() ?? ""
            };
        }

        private static bool HasAdresKorespondencyjnyFields(KontrahentDto kh)
        {
            return kh.Adr_KorespondencyjnyId is > 0
                   || !string.IsNullOrWhiteSpace(kh.Adr_KorespondencyjnyAdres)
                   || !string.IsNullOrWhiteSpace(kh.Adr_KorespondencyjnyUlica)
                   || !string.IsNullOrWhiteSpace(kh.Adr_KorespondencyjnyKod)
                   || !string.IsNullOrWhiteSpace(kh.Adr_KorespondencyjnyMiejscowosc)
                   || !string.IsNullOrWhiteSpace(kh.Adr_KorespondencyjnyNazwa);
        }

        private static string FormatKhTyp(int? typ)
        {
            return typ switch
            {
                0 => "Dostawca i odbiorca",
                1 => "Dostawca",
                2 => "Odbiorca",
                _ => typ.HasValue ? typ.Value.ToString() : ""
            };
        }

        public static string NormalizeDocumentType(string? documentType)
        {
            var typ = (documentType ?? "zk").Trim().ToLowerInvariant();
            return typ is "zk" or "wz" or "fs" ? typ : "zk";
        }

        private static SubiektDocument MapDocument(DocumentDto row, string typ)
        {
            var odbiorca = row.Kh__Kontrahent_Odbiorca ?? row.Kh__Kontrahent_Platnik;
            var platnik = row.Kh__Kontrahent_Platnik;

            var doc = new SubiektDocument
            {
                DokId = row.Dok_Id,
                DokTyp = row.Dok_Typ,
                TypKod = typ,
                Przesylka = MapPrzesylka(row.Pw_Przesylka),
                NrPelny = row.Dok_NrPelny?.Trim() ?? "",
                NrPelnyOryg = row.Dok_NrPelnyOryg?.Trim() ?? "",
                DoDokId = row.Dok_DoDokId is > 0 ? row.Dok_DoDokId : null,
                DoDokNrPelny = SubiektDocumentNumber.Sanitize(row.Dok_DoDokNrPelny),
                DoDokDataWyst = row.Dok_DoDokDataWyst,
                DataWyst = row.Dok_DataWyst,
                WartNetto = row.Dok_WartNetto,
                WartBrutto = row.Dok_WartBrutto,
                StatusNazwa = row.Dok_StatusNazwa?.Trim() ?? "",
                Wystawil = row.Dok_Wystawil?.Trim() ?? "",
                PlatNazwa = row.Dok_PlatNazwa?.Trim() ?? "",
                KartaNazwa = row.Dok_KartaNazwa?.Trim() ?? "",
                KontrahentId = odbiorca?.Kh_Id ?? 0,
                KontrahentSymbol = odbiorca?.Kh_Symbol?.Trim() ?? "",
                KontrahentNazwa = FormatKontrahentNazwa(odbiorca),
                KontrahentNip = odbiorca?.Adr_NIP?.Trim() ?? "",
                KontrahentEmail = odbiorca?.Kh_EMail?.Trim() ?? "",
                KontrahentAdres = FormatKontrahentAdres(odbiorca),
                KontrahentUlica = FormatKontrahentUlica(odbiorca),
                KontrahentKodPocztowy = odbiorca?.Adr_Kod?.Trim() ?? "",
                KontrahentMiejscowosc = odbiorca?.Adr_Miejscowosc?.Trim() ?? "",
                KontrahentKrajKod = FormatKontrahentKrajKod(odbiorca),
                KontrahentTelefon = FormatKontrahentTelefon(odbiorca),
                PlatnikId = platnik?.Kh_Id ?? 0,
                PlatnikNazwa = FormatKontrahentNazwa(platnik),
                PlatnikNip = platnik?.Adr_NIP?.Trim() ?? "",
                PlatnikEmail = platnik?.Kh_EMail?.Trim() ?? "",
                PlatnikAdres = FormatKontrahentAdres(platnik)
            };

            if (row.Dok_Pozycja != null)
            {
                foreach (var line in row.Dok_Pozycja)
                {
                    doc.Pozycje.Add(new SubiektDocumentLine
                    {
                        Plu = line.Tw_PLU > 0 ? line.Tw_PLU : line.Ob_TowId,
                        Symbol = line.Tw_Symbol?.Trim() ?? "",
                        Nazwa = line.Tw_Nazwa?.Trim() ?? "",
                        Ilosc = line.Ob_Ilosc,
                        CenaNetto = line.Ob_CenaNetto,
                        CenaBrutto = line.Ob_CenaBrutto
                    });
                }
            }

            return doc;
        }

        public async Task<SubiektRelatedDocuments?> GetRelatedDocumentsAsync(
            int dokId,
            string? documentType = null,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            if (dokId <= 0)
            {
                return null;
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            var typ = string.IsNullOrWhiteSpace(documentType) ? null : NormalizeDocumentType(documentType);
            var url = string.IsNullOrWhiteSpace(typ)
                ? $"documents/{dokId}/related"
                : $"documents/{typ}/{dokId}/related";

            using var client = CreateClient(config);
            using var response = await client.GetAsync(url, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            var envelope = JsonSerializer.Deserialize<ApiEnvelope<RelatedDataDto>>(body, JsonOptions);
            var data = envelope?.Data;
            if (data == null)
            {
                return null;
            }

            return new SubiektRelatedDocuments
            {
                Zrodlowy = MapRelated(data.Zrodlowy),
                Pochodne = (data.Pochodne ?? new List<RelatedDocDto>())
                    .Select(MapRelated)
                    .Where(d => d != null)
                    .Cast<SubiektRelatedDocument>()
                    .ToList()
            };
        }

        /// <summary>
        /// Related API zwraca tylko 1 poziom (źródłowy + bezpośrednie pochodne).
        /// Dla łańcucha ZK→FS→WZ dociągamy jeszcze jeden poziom, żeby np. na WZ widać było FS i ZK.
        /// </summary>
        public async Task<List<SubiektRelatedDocument>> GetRelatedDocumentsExpandedAsync(
            int dokId,
            string? documentType = null,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            var related = await GetRelatedDocumentsAsync(dokId, documentType, config, cancellationToken)
                .ConfigureAwait(false);
            var items = related?.ToDisplayList() ?? new List<SubiektRelatedDocument>();
            if (items.Count == 0)
            {
                return items;
            }

            config ??= _configService.LoadSubiektConfig();
            var seen = items.Select(d => d.DokId).ToHashSet();
            seen.Add(dokId);

            foreach (var item in items.ToList())
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var nested = await GetRelatedDocumentsAsync(
                        item.DokId, item.TypKod, config, cancellationToken).ConfigureAwait(false);
                    if (nested == null)
                    {
                        continue;
                    }

                    foreach (var extra in nested.ToDisplayList())
                    {
                        if (seen.Add(extra.DokId))
                        {
                            items.Add(extra);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Warning(
                        $"Nie udało się dociągnąć powiązań 2. poziomu dla {item.NrPelny}: {ex.Message}",
                        "SubiektApiService");
                }
            }

            return SortRelatedForDisplay(items);
        }

        private static List<SubiektRelatedDocument> SortRelatedForDisplay(List<SubiektRelatedDocument> items) =>
            items
                .OrderBy(d => d.DokTyp switch
                {
                    SubiektApiDocumentTypes.Zk => 0,
                    SubiektApiDocumentTypes.Wz => 1,
                    SubiektApiDocumentTypes.Fs => 2,
                    _ => 9
                })
                .ThenBy(d => d.NrPelny, StringComparer.OrdinalIgnoreCase)
                .ToList();

        public async Task<(int? FsId, SubiektPrzesylka? Przesylka)> ResolveInvoiceShipmentAsync(
            SubiektDocument document,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            if (document == null || document.DokId <= 0)
            {
                return (null, null);
            }

            config ??= _configService.LoadSubiektConfig();

            if (string.Equals(document.TypKod, "fs", StringComparison.OrdinalIgnoreCase)
                || document.DokTyp == SubiektApiDocumentTypes.Fs)
            {
                return (document.DokId, document.Przesylka);
            }

            try
            {
                var related = await GetRelatedDocumentsAsync(document.DokId, document.TypKod, config, cancellationToken)
                    .ConfigureAwait(false);
                var fs = FindFs(related);

                if (fs == null)
                {
                    var wz = FindWz(related);
                    if (wz != null)
                    {
                        var wzRelated = await GetRelatedDocumentsAsync(wz.DokId, "wz", config, cancellationToken)
                            .ConfigureAwait(false);
                        fs = FindFs(wzRelated);
                    }
                }

                if (fs == null)
                {
                    return (null, null);
                }

                var fsDoc = await GetDocumentByIdAsync("fs", fs.DokId, config, cancellationToken).ConfigureAwait(false);
                return (fs.DokId, fsDoc?.Przesylka);
            }
            catch (Exception ex)
            {
                Warning($"Nie udało się odczytać powiązanego FS dla {document.NrPelny}: {ex.Message}", "SubiektApiService");
                return (null, null);
            }
        }

        public async Task<SubiektPrzesylka?> PutPrzesylkaAsync(
            int fsDokId,
            SubiektPrzesylka przesylka,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            if (fsDokId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(fsDokId));
            }

            if (przesylka == null)
            {
                throw new ArgumentNullException(nameof(przesylka));
            }

            var body = new Dictionary<string, object?>
            {
                ["typ"] = string.IsNullOrWhiteSpace(przesylka.Typ) ? "GLS" : przesylka.Typ.Trim().ToUpperInvariant(),
                ["id"] = przesylka.Id,
                ["nr_listu_przygotowalnia"] = przesylka.NrListuPrzygotowalnia ?? "",
                ["nr_listu_nadane"] = przesylka.NrListuNadane ?? "",
                ["status"] = string.IsNullOrWhiteSpace(przesylka.Status)
                    ? SubiektPrzesylka.StatusUtworzono
                    : przesylka.Status,
                ["data"] = SubiektPrzesylka.FormatData(przesylka.Data ?? DateTime.Now)
            };

            var envelope = await PutPrzesylkaJsonAsync(fsDokId, body, config, cancellationToken).ConfigureAwait(false);
            return envelope?.Data == null ? przesylka : MapPrzesylka(envelope.Data.Pw_Przesylka) ?? przesylka;
        }

        public async Task<SubiektPrzesylka?> ClearPrzesylkaAsync(
            int fsDokId,
            SubiektConfig? config = null,
            CancellationToken cancellationToken = default)
        {
            var now = DateTime.Now;
            var bodyObj = new
            {
                typ = "GLS",
                status = SubiektPrzesylka.StatusUsuniete,
                data = SubiektPrzesylka.FormatData(now)
            };

            var envelope = await PutPrzesylkaJsonAsync(fsDokId, bodyObj, config, cancellationToken).ConfigureAwait(false);
            return envelope?.Data == null
                ? new SubiektPrzesylka
                {
                    Typ = "GLS",
                    Status = SubiektPrzesylka.StatusUsuniete,
                    Data = now
                }
                : MapPrzesylka(envelope.Data.Pw_Przesylka) ?? new SubiektPrzesylka
                {
                    Typ = "GLS",
                    Status = SubiektPrzesylka.StatusUsuniete,
                    Data = now
                };
        }

        private async Task<ApiEnvelope<DocumentDto>?> PutPrzesylkaJsonAsync(
            int fsDokId,
            object bodyObj,
            SubiektConfig? config,
            CancellationToken cancellationToken)
        {
            if (fsDokId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(fsDokId));
            }

            config ??= _configService.LoadSubiektConfig();
            EnsureConfigured(config);

            using var client = CreateClient(config);
            using var content = new StringContent(
                JsonSerializer.Serialize(bodyObj),
                Encoding.UTF8,
                "application/json");

            using var response = await client.PutAsync($"documents/fs/{fsDokId}/przesylka", content, cancellationToken)
                .ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            EnsureSuccess(response, body);

            return JsonSerializer.Deserialize<ApiEnvelope<DocumentDto>>(body, JsonOptions);
        }

        private static SubiektRelatedDocument? FindFs(SubiektRelatedDocuments? related)
        {
            if (related == null)
            {
                return null;
            }

            return related.Pochodne.FirstOrDefault(d => d.IsFs)
                ?? (related.Zrodlowy is { IsFs: true } src ? src : null);
        }

        private static SubiektRelatedDocument? FindWz(SubiektRelatedDocuments? related)
        {
            if (related == null)
            {
                return null;
            }

            return related.Pochodne.FirstOrDefault(d => d.IsWz)
                ?? (related.Zrodlowy is { IsWz: true } src ? src : null);
        }

        private static SubiektRelatedDocument? MapRelated(RelatedDocDto? row)
        {
            if (row == null || row.Dok_Id <= 0)
            {
                return null;
            }

            return new SubiektRelatedDocument
            {
                DokId = row.Dok_Id,
                DokTyp = row.Dok_Typ,
                NrPelny = row.Dok_NrPelny?.Trim() ?? ""
            };
        }

        private static SubiektPrzesylka? MapPrzesylka(JsonElement element)
        {
            if (element.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
            {
                return null;
            }

            JsonElement obj = element;
            if (element.ValueKind == JsonValueKind.String)
            {
                var raw = element.GetString();
                if (string.IsNullOrWhiteSpace(raw))
                {
                    return null;
                }

                try
                {
                    using var parsed = JsonDocument.Parse(raw);
                    obj = parsed.RootElement.Clone();
                }
                catch
                {
                    Warning("Pole pw_Przesylka nie jest poprawnym JSON-em — pomijam.", "SubiektApiService");
                    return null;
                }
            }

            if (obj.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            var id = GetJsonInt(obj, "id");
            var typ = GetJsonString(obj, "typ") ?? "";
            var status = GetJsonString(obj, "status") ?? "";
            var nrPrzygotowalnia = GetJsonString(obj, "nr_listu_przygotowalnia") ?? "";
            var nrNadane = GetJsonString(obj, "nr_listu_nadane") ?? "";
            var data = SubiektPrzesylka.ParseData(GetJsonString(obj, "data"));

            if (id is null or <= 0
                && string.IsNullOrWhiteSpace(status)
                && string.IsNullOrWhiteSpace(typ)
                && string.IsNullOrWhiteSpace(nrPrzygotowalnia)
                && string.IsNullOrWhiteSpace(nrNadane))
            {
                return null;
            }

            return new SubiektPrzesylka
            {
                Typ = typ,
                Id = id is > 0 ? id.Value : 0,
                NrListuPrzygotowalnia = nrPrzygotowalnia,
                NrListuNadane = nrNadane,
                Status = status,
                Data = data
            };
        }

        private static string? GetJsonString(JsonElement obj, string name)
        {
            if (!TryGetPropertyIgnoreCase(obj, name, out var prop))
            {
                return null;
            }

            return prop.ValueKind == JsonValueKind.String ? prop.GetString() : prop.ToString();
        }

        private static int? GetJsonInt(JsonElement obj, string name)
        {
            if (!TryGetPropertyIgnoreCase(obj, name, out var prop))
            {
                return null;
            }

            if (prop.ValueKind == JsonValueKind.Number && prop.TryGetInt32(out var number))
            {
                return number;
            }

            if (prop.ValueKind == JsonValueKind.String
                && int.TryParse(prop.GetString(), out var parsed))
            {
                return parsed;
            }

            return null;
        }

        private static bool TryGetPropertyIgnoreCase(JsonElement obj, string name, out JsonElement value)
        {
            if (obj.TryGetProperty(name, out value))
            {
                return true;
            }

            foreach (var prop in obj.EnumerateObject())
            {
                if (prop.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    value = prop.Value;
                    return true;
                }
            }

            value = default;
            return false;
        }

        private static string FormatKontrahentNazwa(KontrahentDto? kh)
        {
            if (kh == null)
            {
                return "";
            }

            var nazwa = !string.IsNullOrWhiteSpace(kh.Adr_NazwaPelna)
                ? kh.Adr_NazwaPelna
                : kh.Adr_Nazwa;
            if (string.IsNullOrWhiteSpace(nazwa))
            {
                return kh.Kh_Symbol?.Trim() ?? "";
            }

            // Lista / TextBlock: jedna linia zamiast CR/LF z adr_NazwaPelna.
            return string.Join(
                " ",
                nazwa.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim())
                    .Where(part => part.Length > 0));
        }

        private static string FormatKontrahentAdres(KontrahentDto? kh)
        {
            if (kh == null)
            {
                return "";
            }

            var parts = new List<string>();
            var street = FormatKontrahentUlica(kh);
            if (!string.IsNullOrWhiteSpace(street))
            {
                parts.Add(street);
            }

            var city = $"{kh.Adr_Kod} {kh.Adr_Miejscowosc}".Trim();
            if (!string.IsNullOrWhiteSpace(city))
            {
                parts.Add(city);
            }

            return string.Join(", ", parts);
        }

        private static string FormatKontrahentUlica(KontrahentDto? kh)
        {
            if (kh == null)
            {
                return "";
            }

            if (!string.IsNullOrWhiteSpace(kh.Adr_Adres))
            {
                return kh.Adr_Adres.Trim();
            }

            var street = kh.Adr_Ulica?.Trim() ?? "";
            if (!string.IsNullOrWhiteSpace(kh.Adr_NrDomu))
            {
                street = string.IsNullOrWhiteSpace(street)
                    ? kh.Adr_NrDomu.Trim()
                    : $"{street} {kh.Adr_NrDomu.Trim()}";
                if (!string.IsNullOrWhiteSpace(kh.Adr_NrLokalu))
                {
                    street += "/" + kh.Adr_NrLokalu.Trim();
                }
            }

            return street;
        }

        private static string FormatKontrahentKrajKod(KontrahentDto? kh)
        {
            if (kh == null)
            {
                return "PL";
            }

            foreach (var raw in new[] { kh.Pa_Kod, kh.Pa_Symbol, kh.Adr_KrajKod, kh.Adr_Kraj })
            {
                var value = raw?.Trim() ?? "";
                if (value.Length == 2 && value.All(char.IsLetter))
                {
                    return value.ToUpperInvariant();
                }
            }

            var name = kh.Pa_Nazwa?.Trim() ?? kh.Adr_Kraj?.Trim() ?? "";
            if (name.Equals("Polska", StringComparison.OrdinalIgnoreCase)
                || name.Equals("Poland", StringComparison.OrdinalIgnoreCase))
            {
                return "PL";
            }

            return string.IsNullOrWhiteSpace(name) ? "PL" : name;
        }

        private static string FormatKontrahentTelefon(KontrahentDto? kh)
        {
            if (kh == null)
            {
                return "";
            }

            return FirstNonEmpty(kh.Kh_TelefonKomorkowy, kh.Kh_Telefon, kh.Adr_Telefon);
        }

        private static string FirstNonEmpty(params string?[] values)
        {
            foreach (var value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return "";
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
                // Wyszukiwanie po nazwie/NIP na liście dokumentów bywa wolne (20–40 s).
                Timeout = TimeSpan.FromSeconds(90)
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
            public int? Kh_Typ { get; set; }
            public string? Adr_Nazwa { get; set; }
            public string? Adr_NazwaPelna { get; set; }
            public string? Adr_NIP { get; set; }
            public string? Adr_Adres { get; set; }
            public string? Adr_Ulica { get; set; }
            public string? Adr_NrDomu { get; set; }
            public string? Adr_NrLokalu { get; set; }
            public string? Adr_Miejscowosc { get; set; }
            public string? Adr_Kod { get; set; }
            public string? Adr_Poczta { get; set; }
            public string? Adr_Telefon { get; set; }
            public string? Adr_Kraj { get; set; }
            public string? Adr_KrajKod { get; set; }
            public string? Adr_Panstwo { get; set; }
            public string? Adr_Wojewodztwo { get; set; }
            public string? Kh_Telefon { get; set; }
            public string? Kh_TelefonKomorkowy { get; set; }
            public string? Pa_Kod { get; set; }
            public string? Pa_Symbol { get; set; }
            public string? Pa_Nazwa { get; set; }

            // Adres korespondencyjny — GET /kontrahenci/{id}, tylko gdy kh_AdresKoresp=1
            public int? Adr_KorespondencyjnyId { get; set; }
            public string? Adr_KorespondencyjnyNazwa { get; set; }
            public string? Adr_KorespondencyjnyNazwaPelna { get; set; }
            public string? Adr_KorespondencyjnyNIP { get; set; }
            public string? Adr_KorespondencyjnyAdres { get; set; }
            public string? Adr_KorespondencyjnyUlica { get; set; }
            public string? Adr_KorespondencyjnyMiejscowosc { get; set; }
            public string? Adr_KorespondencyjnyKod { get; set; }
            public string? Adr_KorespondencyjnyPoczta { get; set; }
            public string? Adr_KorespondencyjnyTelefon { get; set; }
            public string? Adr_KorespondencyjnyPanstwo { get; set; }
            public string? Adr_KorespondencyjnyWojewodztwo { get; set; }

            // Adres dostawy — GET /kontrahenci/{id}, tylko gdy kh_AdresDostawy=1
            public int? Adr_DostawaId { get; set; }
            public string? Adr_DostawaNazwa { get; set; }
            public string? Adr_DostawaNazwaPelna { get; set; }
            public string? Adr_DostawaNIP { get; set; }
            public string? Adr_DostawaAdres { get; set; }
            public string? Adr_DostawaUlica { get; set; }
            public string? Adr_DostawaMiejscowosc { get; set; }
            public string? Adr_DostawaKod { get; set; }
            public string? Adr_DostawaPoczta { get; set; }
            public string? Adr_DostawaTelefon { get; set; }
            public string? Adr_DostawaPanstwo { get; set; }
            public string? Adr_DostawaWojewodztwo { get; set; }
        }

        private class DocumentDto
        {
            public int Dok_Id { get; set; }
            public string? Dok_NrPelny { get; set; }
            public string? Dok_NrPelnyOryg { get; set; }
            public int Dok_Typ { get; set; }
            public int? Dok_DoDokId { get; set; }
            public string? Dok_DoDokNrPelny { get; set; }
            public DateTime? Dok_DoDokDataWyst { get; set; }
            public DateTime? Dok_DataWyst { get; set; }
            public decimal Dok_WartNetto { get; set; }
            public decimal Dok_WartBrutto { get; set; }
            public int Dok_Status { get; set; }
            public string? Dok_StatusNazwa { get; set; }
            public string? Dok_Wystawil { get; set; }
            public string? Dok_PlatNazwa { get; set; }
            public string? Dok_KartaNazwa { get; set; }
            public JsonElement Pw_Przesylka { get; set; }
            public KontrahentDto? Kh__Kontrahent_Odbiorca { get; set; }
            public KontrahentDto? Kh__Kontrahent_Platnik { get; set; }
            public List<DocumentLineDto>? Dok_Pozycja { get; set; }
        }

        private class RelatedDataDto
        {
            public RelatedDocDto? Zrodlowy { get; set; }
            public List<RelatedDocDto>? Pochodne { get; set; }
        }

        private class RelatedDocDto
        {
            public int Dok_Id { get; set; }
            public int Dok_Typ { get; set; }
            public string? Dok_NrPelny { get; set; }
        }

        private class DocumentLineDto
        {
            public int Ob_Id { get; set; }
            public int Ob_TowId { get; set; }
            public int Tw_PLU { get; set; }
            public string? Tw_Symbol { get; set; }
            public string? Tw_Nazwa { get; set; }
            public decimal Ob_Ilosc { get; set; }
            public decimal Ob_CenaNetto { get; set; }
            public decimal Ob_CenaBrutto { get; set; }
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
