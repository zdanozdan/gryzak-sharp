using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Gryzak.Models;
using static Gryzak.Services.Logger;

namespace Gryzak.Services
{
    public class GlsLoginResult
    {
        public bool Success { get; set; }
        public string? SessionId { get; set; }
        public string? ErrorMessage { get; set; }
        public string EnvironmentName { get; set; } = "";
        public string ApiUrl { get; set; } = "";
    }

    /// <summary>
    /// Klient GLS ADE-Plus WebAPI 2 (SOAP Document/Literal), zgodnie z:
    /// https://ade-test.gls-poland.com/adeplus/pm1/manuals/webapi2_pl/functions/f_login.htm
    /// https://ade-test.gls-poland.com/adeplus/pm1/manuals/webapi2_pl/chapters/guide.htm
    /// </summary>
    public class GlsService : IDisposable
    {
        private const int MaxCredentialLength = 40;
        private const int GlsIdPageSize = 100;
        private const int PickupLookbackDays = 14;
        private const int MaxPickupsToScan = 500;
        private const int PreparingBoxGetConsignConcurrency = 8;
        private const int PreparingBoxProgressLogStep = 50;
        private const int PickupGetConsignConcurrency = 8;

        private readonly GlsConfig _config;
        private readonly HttpClient _httpClient;

        public GlsService(GlsConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));

            var handler = new HttpClientHandler
            {
                CookieContainer = new CookieContainer(),
                UseCookies = true,
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13
            };

            var timeoutSeconds = Math.Clamp(_config.TimeoutSeconds, 5, 300);
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(timeoutSeconds)
            };
        }

        public async Task<GlsLoginResult> TestConnectionAsync(CancellationToken cancellationToken = default)
        {
            var loginResult = await LoginAsync(cancellationToken);
            if (!loginResult.Success || string.IsNullOrWhiteSpace(loginResult.SessionId))
            {
                return loginResult;
            }

            try
            {
                await LogoutAsync(loginResult.SessionId, cancellationToken);
            }
            catch (Exception ex)
            {
                Warning($"Logowanie GLS udane, ale wylogowanie nie powiodło się: {ex.Message}", "GlsService");
            }

            return loginResult;
        }

        public async Task<GlsLoginResult> LoginAsync(CancellationToken cancellationToken = default)
        {
            var result = new GlsLoginResult
            {
                EnvironmentName = _config.GetEnvironmentName(),
                ApiUrl = _config.GetActiveApiUrl()
            };

            if (string.IsNullOrWhiteSpace(_config.UserName) || string.IsNullOrWhiteSpace(_config.Password))
            {
                result.ErrorMessage = "Login i hasło GLS są wymagane.";
                return result;
            }

            if (_config.UserName.Length > MaxCredentialLength || _config.Password.Length > MaxCredentialLength)
            {
                result.ErrorMessage = $"Login i hasło GLS mogą mieć maksymalnie {MaxCredentialLength} znaków (wymaganie WebAPI 2).";
                return result;
            }

            if (string.IsNullOrWhiteSpace(result.ApiUrl))
            {
                result.ErrorMessage = "URL API GLS jest wymagany.";
                return result;
            }

            try
            {
                Info($"Logowanie do GLS ADE ({result.EnvironmentName}): {result.ApiUrl}", "GlsService");

                var responseXml = await SendSoapRequestAsync(
                    "adeLogin",
                    cancellationToken,
                    ("user_name", _config.UserName),
                    ("user_password", _config.Password));

                var fault = GetSoapFault(responseXml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    result.ErrorMessage = fault;
                    Warning($"Logowanie GLS nieudane: {result.ErrorMessage}", "GlsService");
                    return result;
                }

                var sessionId = GetSessionId(responseXml);
                if (string.IsNullOrWhiteSpace(sessionId))
                {
                    result.ErrorMessage = "API GLS nie zwróciło identyfikatora sesji (return.session).";
                    Warning(result.ErrorMessage, "GlsService");
                    return result;
                }

                result.Success = true;
                result.SessionId = sessionId;
                Info("Logowanie do GLS ADE zakończone pomyślnie.", "GlsService");
                return result;
            }
            catch (TaskCanceledException)
            {
                result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                Error(result.ErrorMessage, "GlsService");
                return result;
            }
            catch (HttpRequestException ex)
            {
                result.ErrorMessage = $"Błąd sieci podczas połączenia z GLS: {ex.Message}";
                Error(ex, "GlsService", "Błąd sieci GLS");
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Błąd logowania do GLS: {ex.Message}";
                Error(ex, "GlsService", "Błąd logowania GLS");
                return result;
            }
        }

        public async Task LogoutAsync(string sessionId, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                return;
            }

            await SendSoapRequestAsync("adeLogout", cancellationToken, ("session", sessionId));
            Info("Wylogowano z sesji GLS ADE.", "GlsService");
        }

        /// <summary>
        /// Dodaje przesyłkę do przygotowalni GLS (adePreparingBox_Insert).
        /// https://ade-test.gls-poland.com/adeplus/pm1/html/webapi/functions/f_prepbox_insert.htm
        /// </summary>
        public async Task<GlsCreateShipmentResult> CreateShipmentAsync(
            GlsConsignment consignment,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsCreateShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (consignment == null)
            {
                result.ErrorMessage = "Brak danych przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    var insert = await InsertPreparingBoxAsync(sessionId, consignment, cancellationToken);
                    if (!insert.Success)
                    {
                        result.ErrorMessage = insert.ErrorMessage;
                        return result;
                    }

                    result.Success = true;
                    result.ConsignmentId = insert.ConsignmentId;
                    return result;
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        public async Task<GlsGetShipmentResult> GetShipmentAsync(
            int consignmentId,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsGetShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (consignmentId <= 0)
            {
                result.ErrorMessage = "Brak identyfikatora przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        Info($"GLS adePreparingBox_GetConsign ({result.EnvironmentName}): id={consignmentId}", "GlsService");

                        var bodyXml = BuildIdBody(sessionId, consignmentId);
                        var responseXml = await SendSoapBodyAsync("adePreparingBox_GetConsign", bodyXml, cancellationToken);

                        var fault = GetSoapFault(responseXml);
                        if (!string.IsNullOrWhiteSpace(fault))
                        {
                            result.ErrorMessage = fault;
                            result.NotFound = IsMissingConsignmentFault(fault, responseXml);
                            Warning($"Pobranie przesyłki GLS nieudane: {fault}", "GlsService");
                            return result;
                        }

                        var consignment = ParseConsignment(responseXml);
                        if (consignment == null)
                        {
                            result.ErrorMessage = "API GLS nie zwróciło danych przesyłki.";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        consignment.ExistingId = consignmentId;
                        result.Success = true;
                        result.Consignment = consignment;
                        return result;
                    }
                    catch (TaskCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania przesyłki GLS: {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania przesyłki GLS");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Pobiera dane przesyłki z potwierdzenia nadania (po pickupie).
        /// Jest analogiczne do adePreparingBox_GetConsign, ale dla adePickup_GetConsign.
        /// </summary>
        public async Task<GlsGetShipmentResult> GetPickupShipmentAsync(
            int consignmentId,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsGetShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (consignmentId <= 0)
            {
                result.ErrorMessage = "Brak identyfikatora przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        Info($"GLS adePickup_GetConsign ({result.EnvironmentName}): id={consignmentId}", "GlsService");

                        var bodyXml = BuildIdBody(sessionId, consignmentId);
                        var responseXml = await SendSoapBodyAsync("adePickup_GetConsign", bodyXml, cancellationToken);

                        var fault = GetSoapFault(responseXml);
                        if (!string.IsNullOrWhiteSpace(fault))
                        {
                            result.ErrorMessage = fault;
                            result.NotFound = IsMissingConsignmentFault(fault, responseXml);
                            Warning($"Pobranie przesyłki GLS (pickup) nieudane: {fault}", "GlsService");
                            return result;
                        }

                        var consignment = ParseConsignment(responseXml);
                        if (consignment == null)
                        {
                            result.ErrorMessage = "API GLS nie zwróciło danych przesyłki.";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        consignment.ExistingId = consignmentId;
                        result.Success = true;
                        result.Consignment = consignment;
                        return result;
                    }
                    catch (TaskCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania przesyłki GLS (pickup): {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania przesyłki GLS (pickup)");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Wyszukuje konsygnację po numerze paczki (tracking) w systemie GLS.
        /// Zwykle zwraca paczki z kontekstu „pickup”, ale UI GLS może pokazywać także paczki widoczne wcześniej.
        /// </summary>
        public async Task<GlsGetShipmentResult> SearchParcelNumberAsync(
            string parcelNumber,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsGetShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (string.IsNullOrWhiteSpace(parcelNumber))
            {
                result.ErrorMessage = "Brak numeru paczki.";
                return result;
            }

            parcelNumber = parcelNumber.Trim();

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        Info($"GLS adePickup_ParcelNumberSearch ({result.EnvironmentName}): number={parcelNumber}", "GlsService");

                        var bodyXml = BuildParcelSearchBody(sessionId, parcelNumber);
                        var responseXml = await SendSoapBodyAsync("adePickup_ParcelNumberSearch", bodyXml, cancellationToken);

                        var fault = GetSoapFault(responseXml);
                        if (!string.IsNullOrWhiteSpace(fault))
                        {
                            result.ErrorMessage = fault;
                            result.NotFound = IsMissingConsignmentFault(fault, responseXml);
                            Warning($"Pobranie przesyłki GLS po numerze paczki nieudane: {fault}", "GlsService");
                            return result;
                        }

                        var consignment = ParseConsignment(responseXml);
                        if (consignment == null)
                        {
                            result.ErrorMessage = "API GLS nie zwróciło danych przesyłki.";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        consignment.ExistingId = GetReturnId(responseXml);
                        result.Success = true;
                        result.Consignment = consignment;
                        return result;
                    }
                    catch (TaskCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd wyszukiwania GLS po numerze paczki: {ex.Message}";
                        Error(ex, "GlsService", "Błąd wyszukiwania GLS po numerze paczki");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Tylko przygotowalnia w jednej sesji SOAP.
        /// GLS nie ma wyszukiwania po references / numerze FS — listy zbieramy jak w ADE:
        /// GetIDs (strony po 100) i GetConsign per id (równolegle).
        /// </summary>
        public async Task<GlsPreparingBoxListResult> GetPreparingBoxListAsync(
            CancellationToken cancellationToken = default,
            IProgress<GlsFetchProgress>? progress = null)
        {
            var result = new GlsPreparingBoxListResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        progress?.Report(new GlsFetchProgress
                        {
                            Operation = "GLS przygotowalnia",
                            Phase = "pobieranie listy ID"
                        });
                        result.Items = await FetchPreparingBoxItemsAsync(
                            sessionId, cancellationToken, progress);
                        result.Success = true;
                        Info(
                            $"GLS przygotowalnia: {result.Items.Count} przesyłek ({result.EnvironmentName}).",
                            "GlsService");
                        return result;
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex) when (ex is TaskCanceledException or OperationCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania przygotowalni GLS: {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania przygotowalni GLS");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Tylko potwierdzenia nadania w jednej sesji SOAP, filtrowane po dacie pickup.
        /// </summary>
        public async Task<GlsPickupListResult> GetPickupListAsync(
            GlsPickupQuery query,
            CancellationToken cancellationToken = default,
            IProgress<GlsFetchProgress>? progress = null)
        {
            query ??= GlsPickupQuery.ForLookbackDays(PickupLookbackDays);
            query.Normalize();

            var result = new GlsPickupListResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        progress?.Report(new GlsFetchProgress
                        {
                            Operation = "GLS nadania",
                            Phase = "skanowanie potwierdzeń"
                        });
                        var fetch = await FetchPickupItemsAsync(
                            sessionId, query, cancellationToken, progress);
                        result.Items = fetch.Items;
                        result.HitPickupScanLimit = fetch.HitScanLimit;
                        result.Success = true;
                        Info(
                            $"GLS nadania: {result.Items.Count} przesyłek ({query.FromDate:yyyy-MM-dd}…{query.ToDate:yyyy-MM-dd}, {result.EnvironmentName}).",
                            "GlsService");
                        return result;
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex) when (ex is TaskCanceledException or OperationCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania nadań GLS: {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania nadań GLS");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Przygotowalnia + potwierdzenia nadania w jednej sesji.
        /// GLS nie ma wyszukiwania po references / numerze FS — jedyne szukanie to
        /// adePickup_ParcelNumberSearch po numerze paczki. Dlatego listy zbieramy jak w ADE:
        /// GetIDs (strony po 100) i GetConsign per id.
        /// Nadania: najnowsze pickupy, z datetime nie starszym niż 14 dni.
        /// </summary>
        public async Task<GlsStatusListsResult> GetStatusListsAsync(
            CancellationToken cancellationToken = default)
        {
            var result = new GlsStatusListsResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        result.PreparingBoxItems = await FetchPreparingBoxItemsAsync(sessionId, cancellationToken);
                        try
                        {
                            var pickupQuery = GlsPickupQuery.ForLookbackDays(PickupLookbackDays);
                            var fetch = await FetchPickupItemsAsync(sessionId, pickupQuery, cancellationToken);
                            result.PickupItems = fetch.Items;
                        }
                        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                        {
                            throw;
                        }
                        catch (Exception ex) when (ex is TaskCanceledException or OperationCanceledException)
                        {
                            result.PickupErrorMessage = "Timeout przy pobieraniu nadań GLS.";
                            Warning(result.PickupErrorMessage, "GlsService");
                        }
                        catch (Exception ex)
                        {
                            result.PickupErrorMessage = $"Nie udało się pobrać nadań GLS: {ex.Message}";
                            Warning(result.PickupErrorMessage, "GlsService");
                        }

                        result.Success = true;
                        Info(
                            $"GLS listy: przygotowalnia={result.PreparingBoxItems.Count}, nadania={result.PickupItems.Count}.",
                            "GlsService");
                        return result;
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex) when (ex is TaskCanceledException or OperationCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania list GLS: {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania list GLS");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// Pobiera etykietę PDF/ZPL — najpierw z potwierdzenia nadania, potem z przygotowalni.
        /// Przesyłka bez numerów paczek automatycznie je otrzymuje.
        /// </summary>
        public async Task<GlsLabelResult> GetLabelsAsync(
            int consignmentId,
            string mode = GlsLabelResult.DefaultMode,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsLabelResult
            {
                EnvironmentName = _config.GetEnvironmentName(),
                Mode = string.IsNullOrWhiteSpace(mode) ? GlsLabelResult.DefaultMode : mode.Trim()
            };
            result.FileExtension = GuessLabelExtension(result.Mode);

            if (consignmentId <= 0)
            {
                result.ErrorMessage = "Brak identyfikatora przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    try
                    {
                        var pickupResult = await TryGetPickupLabelsAsync(
                            sessionId, consignmentId, result.Mode, cancellationToken);
                        if (pickupResult != null)
                        {
                            return pickupResult;
                        }

                        Info(
                            $"GLS adePreparingBox_GetConsignLabels ({result.EnvironmentName}): id={consignmentId}, mode={result.Mode}",
                            "GlsService");

                        var bodyXml = BuildLabelsBody(sessionId, consignmentId, result.Mode);
                        var responseXml = await SendSoapBodyAsync("adePreparingBox_GetConsignLabels", bodyXml, cancellationToken);

                        var fault = GetSoapFault(responseXml);
                        if (!string.IsNullOrWhiteSpace(fault))
                        {
                            result.ErrorMessage = fault;
                            result.NotFound = IsMissingConsignmentFault(fault, responseXml);
                            Warning($"Pobranie etykiety GLS nieudane: {fault}", "GlsService");
                            return result;
                        }

                        var labelsBase64 = GetReturnString(responseXml, "labels");
                        if (string.IsNullOrWhiteSpace(labelsBase64))
                        {
                            result.ErrorMessage = "API GLS nie zwróciło pliku etykiety.";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        byte[] bytes;
                        try
                        {
                            bytes = Convert.FromBase64String(labelsBase64.Trim());
                        }
                        catch (FormatException)
                        {
                            result.ErrorMessage = "API GLS zwróciło niepoprawny plik etykiety (base64).";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        if (bytes.Length == 0)
                        {
                            result.ErrorMessage = "API GLS zwróciło pustą etykietę.";
                            Warning(result.ErrorMessage, "GlsService");
                            return result;
                        }

                        result.Success = true;
                        result.LabelBytes = bytes;
                        Info($"Pobrano etykietę GLS ({bytes.Length} B, {result.FileExtension}).", "GlsService");
                        return result;
                    }
                    catch (TaskCanceledException)
                    {
                        result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                        Error(result.ErrorMessage, "GlsService");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Błąd pobierania etykiety GLS: {ex.Message}";
                        Error(ex, "GlsService", "Błąd pobierania etykiety GLS");
                        return result;
                    }
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        /// <summary>
        /// GLS nie ma update — zapis to nowy insert, potem usunięcie starej przesyłki z przygotowalni.
        /// </summary>
        public async Task<GlsCreateShipmentResult> UpdateShipmentAsync(
            int existingId,
            GlsConsignment consignment,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsCreateShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (existingId <= 0)
            {
                result.ErrorMessage = "Brak identyfikatora istniejącej przesyłki.";
                return result;
            }

            if (consignment == null)
            {
                result.ErrorMessage = "Brak danych przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    var insert = await InsertPreparingBoxAsync(sessionId, consignment, cancellationToken);
                    if (!insert.Success || insert.ConsignmentId is not > 0)
                    {
                        result.ErrorMessage = insert.ErrorMessage ?? "Nie udało się zapisać listu GLS.";
                        return result;
                    }

                    result.Success = true;
                    result.ConsignmentId = insert.ConsignmentId;

                    var delete = await DeletePreparingBoxAsync(sessionId, existingId, cancellationToken);
                    if (!delete.Success)
                    {
                        result.WarningMessage =
                            $"Zapisano nową przesyłkę (id={insert.ConsignmentId}), ale nie usunięto poprzedniej (id={existingId}): {delete.ErrorMessage}";
                        Warning(result.WarningMessage, "GlsService");
                    }

                    return result;
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        public async Task<GlsCreateShipmentResult> DeleteShipmentAsync(
            int consignmentId,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsCreateShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            if (consignmentId <= 0)
            {
                result.ErrorMessage = "Brak identyfikatora przesyłki.";
                return result;
            }

            return await WithSessionAsync(
                async sessionId =>
                {
                    var deleted = await DeletePreparingBoxAsync(sessionId, consignmentId, cancellationToken);
                    result.Success = deleted.Success;
                    result.NotFound = deleted.NotFound;
                    result.ConsignmentId = deleted.ConsignmentId;
                    result.ErrorMessage = deleted.ErrorMessage;
                    return result;
                },
                loginError =>
                {
                    result.ErrorMessage = loginError;
                    return result;
                },
                cancellationToken);
        }

        private async Task<T> WithSessionAsync<T>(
            Func<string, Task<T>> action,
            Func<string, T> onLoginFail,
            CancellationToken cancellationToken)
        {
            var login = await LoginAsync(cancellationToken);
            if (!login.Success || string.IsNullOrWhiteSpace(login.SessionId))
            {
                return onLoginFail(login.ErrorMessage ?? "Nie udało się zalogować do GLS.");
            }

            try
            {
                return await action(login.SessionId);
            }
            finally
            {
                try
                {
                    await LogoutAsync(login.SessionId, cancellationToken);
                }
                catch (Exception ex)
                {
                    Warning($"Wylogowanie GLS nie powiodło się: {ex.Message}", "GlsService");
                }
            }
        }

        private async Task<GlsCreateShipmentResult> InsertPreparingBoxAsync(
            string sessionId,
            GlsConsignment consignment,
            CancellationToken cancellationToken)
        {
            var result = new GlsCreateShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            try
            {
                Info(
                    $"GLS adePreparingBox_Insert ({result.EnvironmentName}): {consignment.Name1}, {consignment.SummaryAddress}",
                    "GlsService");

                var bodyXml = BuildPreparingBoxInsertBody(sessionId, consignment);
                var responseXml = await SendSoapBodyAsync("adePreparingBox_Insert", bodyXml, cancellationToken);

                var fault = GetSoapFault(responseXml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    result.ErrorMessage = fault;
                    Warning($"Tworzenie przesyłki GLS nieudane: {fault}", "GlsService");
                    return result;
                }

                var id = GetReturnId(responseXml);
                if (id == null)
                {
                    result.ErrorMessage = "API GLS nie zwróciło identyfikatora przesyłki (return.id).";
                    Warning(result.ErrorMessage, "GlsService");
                    return result;
                }

                result.Success = true;
                result.ConsignmentId = id;
                Info($"Utworzono przesyłkę GLS w przygotowalni, id={id}.", "GlsService");
                return result;
            }
            catch (TaskCanceledException)
            {
                result.ErrorMessage = "Timeout — połączenie z GLS przekroczyło limit czasu.";
                Error(result.ErrorMessage, "GlsService");
                return result;
            }
            catch (HttpRequestException ex)
            {
                result.ErrorMessage = $"Błąd sieci podczas połączenia z GLS: {ex.Message}";
                Error(ex, "GlsService", "Błąd sieci GLS");
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Błąd tworzenia przesyłki GLS: {ex.Message}";
                Error(ex, "GlsService", "Błąd tworzenia przesyłki GLS");
                return result;
            }
        }

        private async Task<GlsCreateShipmentResult> DeletePreparingBoxAsync(
            string sessionId,
            int consignmentId,
            CancellationToken cancellationToken)
        {
            var result = new GlsCreateShipmentResult
            {
                EnvironmentName = _config.GetEnvironmentName()
            };

            try
            {
                Info($"GLS adePreparingBox_DeleteConsign ({result.EnvironmentName}): id={consignmentId}", "GlsService");

                var bodyXml = BuildIdBody(sessionId, consignmentId);
                var responseXml = await SendSoapBodyAsync("adePreparingBox_DeleteConsign", bodyXml, cancellationToken);

                var fault = GetSoapFault(responseXml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    result.ErrorMessage = fault;
                    result.NotFound = IsMissingConsignmentFault(fault, responseXml);
                    Warning($"Usunięcie przesyłki GLS nieudane: {fault}", "GlsService");
                    return result;
                }

                result.Success = true;
                result.ConsignmentId = consignmentId;
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Błąd usuwania przesyłki GLS: {ex.Message}";
                Error(ex, "GlsService", "Błąd usuwania przesyłki GLS");
                return result;
            }
        }

        private async Task<XDocument> SendSoapRequestAsync(
            string operation,
            CancellationToken cancellationToken,
            params (string Name, string Value)[] parameters)
        {
            var inner = new StringBuilder();
            foreach (var (name, value) in parameters)
            {
                AppendSoapString(inner, name, value);
            }

            return await SendSoapBodyAsync(operation, inner.ToString(), cancellationToken);
        }

        private async Task<XDocument> SendSoapBodyAsync(
            string operation,
            string bodyInnerXml,
            CancellationToken cancellationToken)
        {
            var apiUrl = _config.GetActiveApiUrl();
            var soapXml = BuildSoapEnvelope(apiUrl, operation, bodyInnerXml);
            var soapAction = $"{apiUrl.Trim()}#{operation}";

            using var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
            request.Headers.TryAddWithoutValidation("SOAPAction", $"\"{soapAction}\"");
            request.Content = new StringContent(soapXml, Encoding.UTF8);
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("text/xml") { CharSet = "utf-8" };

            Debug($"SOAP {operation} → {apiUrl}", "GlsService");
            Debug($"SOAP żądanie {operation}: {MaskSecrets(soapXml)}", "GlsService");

            using var response = await _httpClient.SendAsync(request, cancellationToken);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken);
            Debug($"SOAP odpowiedź {operation} ({(int)response.StatusCode} {response.ReasonPhrase}): {MaskSecrets(responseText)}", "GlsService");

            if (string.IsNullOrWhiteSpace(responseText))
            {
                throw new Exception($"Pusta odpowiedź SOAP ({(int)response.StatusCode} {response.ReasonPhrase}).");
            }

            try
            {
                return XDocument.Parse(responseText);
            }
            catch (Exception ex)
            {
                throw new Exception($"Nieprawidłowa odpowiedź SOAP GLS ({(int)response.StatusCode}): {ex.Message}");
            }
        }

        /// <summary>
        /// Envelope jak z PHP SoapClient (Document/Literal): operacja w namespace WSDL,
        /// pola user_name / user_password bez namespace — tak GLS serializuje stdClass z dokumentacji.
        /// </summary>
        private static string BuildSoapEnvelope(string apiUrl, string operation, string bodyInnerXml)
        {
            var ns = SecurityElement.Escape(apiUrl.Trim());
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<SOAP-ENV:Envelope xmlns:SOAP-ENV=\"http://schemas.xmlsoap.org/soap/envelope/\"");
            sb.Append(" xmlns:ns1=\"").Append(ns).Append('"');
            sb.Append(" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            sb.Append(" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">");
            sb.Append("<SOAP-ENV:Body>");
            sb.Append("<ns1:").Append(operation).Append('>');
            sb.Append(bodyInnerXml);
            sb.Append("</ns1:").Append(operation).Append('>');
            sb.Append("</SOAP-ENV:Body>");
            sb.Append("</SOAP-ENV:Envelope>");
            return sb.ToString();
        }

        private static string BuildPreparingBoxInsertBody(string sessionId, GlsConsignment consignment)
        {
            var sb = new StringBuilder();
            AppendSoapString(sb, "session", sessionId);
            sb.Append("<consign_prep_data>");
            AppendSoapString(sb, "rname1", consignment.Name1);
            AppendSoapString(sb, "rname2", consignment.Name2, omitIfEmpty: true);
            AppendSoapString(sb, "rname3", consignment.Name3, omitIfEmpty: true);
            AppendSoapString(sb, "rcountry", consignment.Country);
            AppendSoapString(sb, "rzipcode", consignment.ZipCode);
            AppendSoapString(sb, "rcity", consignment.City);
            AppendSoapString(sb, "rstreet", consignment.Street);
            AppendSoapString(sb, "rphone", consignment.Phone, omitIfEmpty: true);
            AppendSoapString(sb, "rcontact", consignment.Contact, omitIfEmpty: true);
            AppendSoapString(sb, "references", consignment.References, omitIfEmpty: true);
            AppendSoapString(sb, "notes", consignment.Notes, omitIfEmpty: true);

            if (consignment.CashOnDelivery)
            {
                sb.Append("<srv_bool>");
                AppendSoapTyped(sb, "cod", "true", "xsd:boolean");
                AppendSoapTyped(
                    sb,
                    "cod_amount",
                    NormalizeCodAmount(consignment.CodAmount),
                    "xsd:float");
                sb.Append("</srv_bool>");
            }

            var parcels = consignment.Parcels != null && consignment.Parcels.Count > 0
                ? consignment.Parcels
                : new List<GlsParcel> { new() { Weight = GlsConsignment.DefaultParcelWeightKg } };

            sb.Append("<parcels>");
            foreach (var parcel in parcels)
            {
                sb.Append("<items>");
                AppendSoapString(sb, "reference", parcel.Reference, omitIfEmpty: true);
                AppendSoapTyped(sb, "weight", NormalizeWeight(parcel.Weight), "xsd:float");
                sb.Append("</items>");
            }
            sb.Append("</parcels>");
            sb.Append("</consign_prep_data>");
            return sb.ToString();
        }

        private static string BuildIdBody(string sessionId, int consignmentId)
        {
            var sb = new StringBuilder();
            AppendSoapString(sb, "session", sessionId);
            AppendSoapTyped(sb, "id", consignmentId.ToString(CultureInfo.InvariantCulture), "xsd:long");
            return sb.ToString();
        }

        private static string BuildParcelSearchBody(string sessionId, string parcelNumber)
        {
            var sb = new StringBuilder();
            AppendSoapString(sb, "session", sessionId);
            AppendSoapTyped(sb, "number", parcelNumber, "xsd:string");
            return sb.ToString();
        }

        private static string BuildIdStartBody(string sessionId, int idStart, int? parentId = null)
        {
            var sb = new StringBuilder();
            AppendSoapString(sb, "session", sessionId);
            if (parentId is > 0)
            {
                AppendSoapTyped(sb, "id", parentId.Value.ToString(CultureInfo.InvariantCulture), "xsd:long");
            }

            AppendSoapTyped(sb, "id_start", idStart.ToString(CultureInfo.InvariantCulture), "xsd:long");
            return sb.ToString();
        }

        private async Task<List<GlsPreparingBoxItem>> FetchPreparingBoxItemsAsync(
            string sessionId,
            CancellationToken cancellationToken,
            IProgress<GlsFetchProgress>? progress = null)
        {
            var ids = await GetPagedIdsAsync(
                "adePreparingBox_GetConsignIDs",
                sessionId,
                cancellationToken);
            Info($"GLS adePreparingBox_GetConsignIDs: {ids.Count} przesyłek", "GlsService");

            if (ids.Count == 0)
            {
                progress?.Report(new GlsFetchProgress
                {
                    Operation = "GLS przygotowalnia",
                    Phase = "brak przesyłek",
                    Done = 0,
                    Total = 0
                });
                return new List<GlsPreparingBoxItem>();
            }

            progress?.Report(new GlsFetchProgress
            {
                Operation = "GLS przygotowalnia",
                Phase = "pobieranie przesyłek",
                Done = 0,
                Total = ids.Count
            });

            var items = new GlsPreparingBoxItem[ids.Count];
            using var gate = new SemaphoreSlim(PreparingBoxGetConsignConcurrency);
            var done = 0;

            var tasks = ids.Select(async (id, index) =>
            {
                await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
                string? refs = null;
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var consignment = await TryGetConsignmentAsync(
                        "adePreparingBox_GetConsign", sessionId, id, cancellationToken)
                        .ConfigureAwait(false);
                    refs = GetConsignmentReferences(consignment);
                    items[index] = new GlsPreparingBoxItem
                    {
                        Id = id,
                        References = refs,
                        ParcelNumber = consignment?.GetParcelNumbersDisplay() ?? ""
                    };
                }
                finally
                {
                    gate.Release();
                    var completed = Interlocked.Increment(ref done);
                    progress?.Report(new GlsFetchProgress
                    {
                        Operation = "GLS przygotowalnia",
                        Phase = "pobieranie przesyłek",
                        Done = completed,
                        Total = ids.Count,
                        CurrentId = id,
                        CurrentLabel = TruncateProgressLabel(refs)
                    });
                    if (completed == ids.Count
                        || completed % PreparingBoxProgressLogStep == 0)
                    {
                        Info(
                            $"GLS przygotowalnia: pobrano {completed}/{ids.Count} GetConsign.",
                            "GlsService");
                    }
                }
            });

            await Task.WhenAll(tasks).ConfigureAwait(false);
            return items.Where(i => i != null).ToList()!;
        }

        private static string? TruncateProgressLabel(string? value)
        {
            var text = (value ?? "").Trim();
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            return text.Length <= 40 ? text : text[..37] + "...";
        }

        private sealed class PickupFetchResult
        {
            public List<GlsPickupItem> Items { get; set; } = new();
            public bool HitScanLimit { get; set; }
        }

        private async Task<PickupFetchResult> FetchPickupItemsAsync(
            string sessionId,
            GlsPickupQuery query,
            CancellationToken cancellationToken,
            IProgress<GlsFetchProgress>? progress = null)
        {
            query.Normalize();
            var fromDate = query.FromDate;
            var toDate = query.ToDate;
            var result = new PickupFetchResult();
            var pickupJobs = new List<(int PickupId, List<int> ConsignIds)>();
            var idStart = 0;
            var reachedOlderThanFrom = false;
            var hitLimit = false;
            var scannedPickups = 0;

            // Faza 1: zbierz pickupi i ID przesyłek w oknie dat.
            while (pickupJobs.Count < MaxPickupsToScan)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var xml = await SendSoapBodyAsync(
                    "adePickup_GetIDs",
                    BuildIdStartBody(sessionId, idStart),
                    cancellationToken);
                var fault = GetSoapFault(xml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    throw new Exception(fault);
                }

                var batch = ParseIdArray(xml);
                if (batch.Count == 0)
                {
                    break;
                }

                foreach (var pickupId in batch)
                {
                    if (pickupJobs.Count >= MaxPickupsToScan)
                    {
                        hitLimit = true;
                        break;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    scannedPickups++;
                    progress?.Report(new GlsFetchProgress
                    {
                        Operation = "GLS nadania",
                        Phase = "skanowanie potwierdzeń",
                        Done = scannedPickups,
                        Total = 0,
                        CurrentId = pickupId
                    });

                    var createdAt = await TryGetPickupCreatedAtAsync(sessionId, pickupId, cancellationToken);
                    if (createdAt is DateTime when)
                    {
                        var day = when.Date;
                        if (day < fromDate)
                        {
                            reachedOlderThanFrom = true;
                            break;
                        }

                        if (day > toDate)
                        {
                            continue;
                        }
                    }

                    var consignIds = await GetPagedIdsAsync(
                        "adePickup_GetConsignIDs",
                        sessionId,
                        cancellationToken,
                        parentId: pickupId);
                    if (consignIds.Count == 0)
                    {
                        continue;
                    }

                    pickupJobs.Add((pickupId, consignIds));
                }

                if (reachedOlderThanFrom || hitLimit || batch.Count < GlsIdPageSize)
                {
                    break;
                }

                var nextStart = batch[^1];
                if (nextStart <= 0 || nextStart == idStart)
                {
                    break;
                }

                idStart = nextStart;
            }

            var consignWork = pickupJobs
                .SelectMany(job => job.ConsignIds.Select(consignId => (job.PickupId, ConsignId: consignId)))
                .ToList();

            if (consignWork.Count == 0)
            {
                progress?.Report(new GlsFetchProgress
                {
                    Operation = "GLS nadania",
                    Phase = "brak przesyłek w zakresie",
                    Done = 0,
                    Total = 0
                });
                result.HitScanLimit = hitLimit;
                return result;
            }

            progress?.Report(new GlsFetchProgress
            {
                Operation = "GLS nadania",
                Phase = "pobieranie przesyłek",
                Done = 0,
                Total = consignWork.Count
            });

            var items = new GlsPickupItem[consignWork.Count];
            using var gate = new SemaphoreSlim(PickupGetConsignConcurrency);
            var done = 0;

            var tasks = consignWork.Select(async (work, index) =>
            {
                await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
                string? refs = null;
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var consignment = await TryGetConsignmentAsync(
                        "adePickup_GetConsign", sessionId, work.ConsignId, cancellationToken)
                        .ConfigureAwait(false);
                    refs = GetConsignmentReferences(consignment);
                    items[index] = new GlsPickupItem
                    {
                        Id = work.ConsignId,
                        PickupId = work.PickupId,
                        References = refs,
                        ParcelNumber = consignment?.GetParcelNumbersDisplay() ?? ""
                    };
                }
                finally
                {
                    gate.Release();
                    var completed = Interlocked.Increment(ref done);
                    progress?.Report(new GlsFetchProgress
                    {
                        Operation = "GLS nadania",
                        Phase = "pobieranie przesyłek",
                        Done = completed,
                        Total = consignWork.Count,
                        CurrentId = work.ConsignId,
                        CurrentLabel = TruncateProgressLabel(refs)
                    });
                }
            });

            await Task.WhenAll(tasks).ConfigureAwait(false);
            result.Items.AddRange(items.Where(i => i != null)!);

            if (hitLimit || pickupJobs.Count >= MaxPickupsToScan)
            {
                result.HitScanLimit = true;
                Warning(
                    $"GLS nadania: osiągnięto limit skanu {MaxPickupsToScan} potwierdzeń (okno {fromDate:yyyy-MM-dd}…{toDate:yyyy-MM-dd}).",
                    "GlsService");
            }

            Info(
                $"GLS nadania: {pickupJobs.Count} potwierdzeń, {result.Items.Count} przesyłek ({fromDate:yyyy-MM-dd}…{toDate:yyyy-MM-dd}).",
                "GlsService");
            return result;
        }

        private async Task<DateTime?> TryGetPickupCreatedAtAsync(
            string sessionId,
            int pickupId,
            CancellationToken cancellationToken)
        {
            try
            {
                var xml = await SendSoapBodyAsync(
                    "adePickup_Get",
                    BuildIdBody(sessionId, pickupId),
                    cancellationToken);
                if (!string.IsNullOrWhiteSpace(GetSoapFault(xml)))
                {
                    return null;
                }

                var text = GetReturnString(xml, "datetime");
                if (string.IsNullOrWhiteSpace(text))
                {
                    return null;
                }

                if (DateTime.TryParseExact(
                        text,
                        "yyyy-MM-dd HH:mm:ss",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out var exact))
                {
                    return exact;
                }

                return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed)
                    ? parsed
                    : null;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Warning($"GLS adePickup_Get id={pickupId}: {ex.Message}", "GlsService");
                return null;
            }
        }

        private async Task<List<int>> GetPagedIdsAsync(
            string operation,
            string sessionId,
            CancellationToken cancellationToken,
            int? parentId = null)
        {
            var ids = new List<int>();
            var seen = new HashSet<int>();
            var idStart = 0;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var xml = await SendSoapBodyAsync(
                    operation,
                    BuildIdStartBody(sessionId, idStart, parentId),
                    cancellationToken);

                var fault = GetSoapFault(xml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    throw new Exception(fault);
                }

                var batch = ParseIdArray(xml);
                if (batch.Count == 0)
                {
                    break;
                }

                foreach (var id in batch)
                {
                    if (seen.Add(id))
                    {
                        ids.Add(id);
                    }
                }

                if (batch.Count < GlsIdPageSize)
                {
                    break;
                }

                var nextStart = batch[^1];
                if (nextStart <= 0 || nextStart == idStart)
                {
                    break;
                }

                idStart = nextStart;
            }

            return ids;
        }

        private async Task<GlsConsignment?> TryGetConsignmentAsync(
            string operation,
            string sessionId,
            int consignmentId,
            CancellationToken cancellationToken)
        {
            try
            {
                var xml = await SendSoapBodyAsync(
                    operation,
                    BuildIdBody(sessionId, consignmentId),
                    cancellationToken);

                var fault = GetSoapFault(xml);
                if (!string.IsNullOrWhiteSpace(fault))
                {
                    Warning($"GLS {operation} id={consignmentId}: {fault}", "GlsService");
                    return null;
                }

                var consignment = ParseConsignment(xml);
                if (consignment == null)
                {
                    var references = GetReturnString(xml, "references") ?? "";
                    if (string.IsNullOrWhiteSpace(references))
                    {
                        return null;
                    }

                    consignment = new GlsConsignment { References = references.Trim() };
                }

                consignment.ExistingId = consignmentId;
                return consignment;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Warning($"GLS {operation} id={consignmentId}: {ex.Message}", "GlsService");
                return null;
            }
        }

        private static string GetConsignmentReferences(GlsConsignment? consignment)
        {
            var references = (consignment?.References ?? "").Trim();
            if (!string.IsNullOrEmpty(references))
            {
                return references;
            }

            return (consignment?.Parcels.FirstOrDefault()?.Reference ?? "").Trim();
        }

        private static List<int> ParseIdArray(XDocument document)
        {
            var ids = new List<int>();
            var returnElement = document.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "return");
            if (returnElement == null)
            {
                return ids;
            }

            foreach (var el in returnElement.Elements())
            {
                if (el.HasElements)
                {
                    foreach (var child in el.Elements())
                    {
                        TryAddParsedId(ids, child.Value);
                    }

                    continue;
                }

                TryAddParsedId(ids, el.Value);
            }

            return ids;
        }

        private static void TryAddParsedId(List<int> ids, string? raw)
        {
            if (int.TryParse((raw ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                && id > 0)
            {
                ids.Add(id);
            }
        }

        private static string BuildLabelsBody(string sessionId, int consignmentId, string mode)
        {
            var sb = new StringBuilder();
            AppendSoapString(sb, "session", sessionId);
            AppendSoapTyped(sb, "id", consignmentId.ToString(CultureInfo.InvariantCulture), "xsd:long");
            AppendSoapString(sb, "mode", mode);
            return sb.ToString();
        }

        private async Task<GlsLabelResult?> TryGetPickupLabelsAsync(
            string sessionId,
            int consignmentId,
            string mode,
            CancellationToken cancellationToken)
        {
            var result = new GlsLabelResult
            {
                EnvironmentName = _config.GetEnvironmentName(),
                Mode = mode,
                FileExtension = GuessLabelExtension(mode)
            };

            Info(
                $"GLS adePickup_GetConsignLabels ({result.EnvironmentName}): id={consignmentId}, mode={mode}",
                "GlsService");

            var bodyXml = BuildLabelsBody(sessionId, consignmentId, mode);
            var responseXml = await SendSoapBodyAsync("adePickup_GetConsignLabels", bodyXml, cancellationToken);
            var fault = GetSoapFault(responseXml);
            if (!string.IsNullOrWhiteSpace(fault))
            {
                if (IsMissingConsignmentFault(fault, responseXml))
                {
                    return null;
                }

                result.ErrorMessage = fault;
                result.NotFound = true;
                Warning($"Pobranie etykiety z pickupu GLS nieudane: {fault}", "GlsService");
                return result;
            }

            var labelsBase64 = GetReturnString(responseXml, "labels");
            if (string.IsNullOrWhiteSpace(labelsBase64))
            {
                return null;
            }

            try
            {
                result.LabelBytes = Convert.FromBase64String(labelsBase64.Trim());
            }
            catch (FormatException ex)
            {
                result.ErrorMessage = $"Nieprawidłowe dane etykiety GLS (base64): {ex.Message}";
                Warning(result.ErrorMessage, "GlsService");
                return result;
            }

            result.Success = true;
            return result;
        }

        private static string GuessLabelExtension(string mode)
        {
            var m = (mode ?? "").Trim().ToLowerInvariant();
            if (m.Contains("zebra_epl") || m.EndsWith("_epl"))
            {
                return ".epl";
            }

            if (m.Contains("zebra") || m.EndsWith("_zpl"))
            {
                return ".zpl";
            }

            if (m.Contains("datamax") || m.EndsWith("_dpl"))
            {
                return ".dpl";
            }

            return ".pdf";
        }

        private static GlsConsignment? ParseConsignment(XDocument document)
        {
            var root = document.Descendants().FirstOrDefault(e => e.Name.LocalName == "return")
                ?? document.Root;
            if (root == null)
            {
                return null;
            }

            var name1 = GetChildValue(root, "rname1");
            var street = GetChildValue(root, "rstreet");
            if (string.IsNullOrWhiteSpace(name1) && string.IsNullOrWhiteSpace(street))
            {
                return null;
            }

            var consignment = new GlsConsignment
            {
                Name1 = name1,
                Name2 = GetChildValue(root, "rname2"),
                Name3 = GetChildValue(root, "rname3"),
                Country = string.IsNullOrWhiteSpace(GetChildValue(root, "rcountry"))
                    ? "PL"
                    : GetChildValue(root, "rcountry"),
                ZipCode = GetChildValue(root, "rzipcode"),
                City = GetChildValue(root, "rcity"),
                Street = street,
                Phone = GetChildValue(root, "rphone"),
                Contact = GetChildValue(root, "rcontact"),
                References = GetChildValue(root, "references"),
                Notes = GetChildValue(root, "notes")
            };

            ApplySrvBool(root, consignment);
            ParseParcelsInto(root, consignment);

            if (consignment.Parcels.Count == 0)
            {
                var weight = GetChildValue(root, "weight");
                var number = GetChildValue(root, "number");
                consignment.Parcels.Add(new GlsParcel
                {
                    Weight = string.IsNullOrWhiteSpace(weight) ? GlsConsignment.DefaultParcelWeightKg : weight,
                    Reference = consignment.References,
                    Number = number
                });
            }

            return consignment;
        }

        private static void ParseParcelsInto(XElement root, GlsConsignment consignment)
        {
            // ADE bywa różne: <parcels><items/>…</parcels> albo wiele <parcels/> / <parcel/>.
            var parcelContainers = root.Elements()
                .Where(e =>
                    e.Name.LocalName.Equals("parcels", StringComparison.OrdinalIgnoreCase)
                    || e.Name.LocalName.Equals("parcel", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var container in parcelContainers)
            {
                var items = container.Elements()
                    .Where(e =>
                        e.Name.LocalName.Equals("items", StringComparison.OrdinalIgnoreCase)
                        || e.Name.LocalName.Equals("item", StringComparison.OrdinalIgnoreCase)
                        || e.Name.LocalName.Equals("parcel", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (items.Count > 0)
                {
                    foreach (var item in items)
                    {
                        AddParsedParcel(consignment, item);
                    }

                    continue;
                }

                // Sam kontener to jedna paczka (weight/number bezpośrednio w <parcels>).
                if (!string.IsNullOrWhiteSpace(GetChildValue(container, "number"))
                    || !string.IsNullOrWhiteSpace(GetChildValue(container, "weight")))
                {
                    AddParsedParcel(consignment, container);
                }
            }
        }

        private static void AddParsedParcel(GlsConsignment consignment, XElement item)
        {
            consignment.Parcels.Add(new GlsParcel
            {
                Weight = GetChildValue(item, "weight"),
                Reference = GetChildValue(item, "reference"),
                Number = GetChildValue(item, "number")
            });
        }

        private static void ApplySrvBool(XElement root, GlsConsignment consignment)
        {
            var srvBool = root.Elements().FirstOrDefault(e =>
                e.Name.LocalName.Equals("srv_bool", StringComparison.OrdinalIgnoreCase));
            if (srvBool == null)
            {
                return;
            }

            consignment.CashOnDelivery = IsTruthy(GetChildValue(srvBool, "cod"));
            var amountText = GetChildValue(srvBool, "cod_amount");
            if (decimal.TryParse(amountText, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount)
                || decimal.TryParse(amountText, NumberStyles.Any, CultureInfo.GetCultureInfo("pl-PL"), out amount))
            {
                consignment.CodAmount = amount;
            }
        }

        private static bool IsTruthy(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var text = value.Trim();
            return text is "1" or "true" or "True" or "TRUE" or "yes" or "Yes";
        }

        private static string GetChildValue(XElement parent, string localName)
        {
            return parent.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                ?.Value?.Trim()
                ?? "";
        }

        private static void AppendSoapString(StringBuilder sb, string name, string? value, bool omitIfEmpty = false)
        {
            if (omitIfEmpty && string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            AppendSoapTyped(sb, name, value ?? "", "xsd:string");
        }

        private static void AppendSoapTyped(StringBuilder sb, string name, string value, string xsdType)
        {
            sb.Append('<').Append(name).Append(" xsi:type=\"").Append(xsdType).Append("\">");
            sb.Append(SecurityElement.Escape(value) ?? "");
            sb.Append("</").Append(name).Append('>');
        }

        private static string NormalizeWeight(string? weight)
        {
            if (decimal.TryParse(weight, NumberStyles.Any, CultureInfo.InvariantCulture, out var kg)
                && kg >= 0.01m)
            {
                return kg.ToString("0.##", CultureInfo.InvariantCulture);
            }

            return GlsConsignment.DefaultParcelWeightKg;
        }

        private static string NormalizeCodAmount(decimal amount)
        {
            var value = amount < 0 ? 0 : amount;
            return value.ToString("0.00", CultureInfo.InvariantCulture);
        }

        private static string? GetSessionId(XDocument document)
        {
            var returnElement = document.Descendants().FirstOrDefault(e => e.Name.LocalName == "return");
            var session = returnElement?.Elements().FirstOrDefault(e => e.Name.LocalName == "session")?.Value?.Trim();
            if (!string.IsNullOrWhiteSpace(session))
            {
                return session;
            }

            return document.Descendants().FirstOrDefault(e => e.Name.LocalName == "session")?.Value?.Trim();
        }

        private static int? GetReturnId(XDocument document)
        {
            var returnElement = document.Descendants().FirstOrDefault(e => e.Name.LocalName == "return");
            var idText = returnElement?.Elements().FirstOrDefault(e => e.Name.LocalName == "id")?.Value?.Trim()
                ?? document.Descendants().FirstOrDefault(e => e.Name.LocalName == "id")?.Value?.Trim();

            return int.TryParse(idText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                ? id
                : null;
        }

        private static string? GetReturnString(XDocument document, string localName)
        {
            var returnElement = document.Descendants().FirstOrDefault(e => e.Name.LocalName == "return");
            var value = returnElement?.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                ?.Value?.Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            return document.Descendants()
                .FirstOrDefault(e => e.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                ?.Value?.Trim();
        }

        private static string? GetSoapFault(XDocument document)
        {
            var fault = document.Descendants().FirstOrDefault(e => e.Name.LocalName == "Fault");
            if (fault == null)
            {
                return null;
            }

            // Dokumentacja GLS: faultcode = kod błędu, faultstring = dopisek (np. "user: 123"),
            // faultactor = nazwa metody, detail = opcjonalne dodatkowe informacje.
            var faultCode = GetDirectChildValue(fault, "faultcode");
            var faultString = GetDirectChildValue(fault, "faultstring")
                ?? fault.Descendants().FirstOrDefault(e => e.Name.LocalName == "Text")?.Value?.Trim();
            var faultActor = GetDirectChildValue(fault, "faultactor");
            var detail = GetFaultDetail(fault);

            Debug($"SOAP Fault code={faultCode}; string={faultString}; actor={faultActor}; detail={detail}", "GlsService");

            var errorCode = ExtractErrorCode(faultCode)
                ?? ExtractErrorCode(faultString)
                ?? ExtractErrorCode(detail)
                ?? ExtractErrorCode(fault.Value);

            var extras = new[] { faultString, detail }
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!.Trim())
                .Where(s => errorCode == null || !s.Equals(errorCode, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var message = !string.IsNullOrWhiteSpace(errorCode)
                ? TranslateFault(errorCode)
                : string.IsNullOrWhiteSpace(faultCode) || IsGenericSoapCode(faultCode)
                    ? "Błąd GLS."
                    : $"Błąd GLS: {faultCode}.";

            if (extras.Count > 0)
            {
                message += " Szczegóły API: " + string.Join(" | ", extras);
            }

            if (!string.IsNullOrWhiteSpace(faultActor))
            {
                message += $" (metoda: {faultActor})";
            }

            return message;
        }

        private static string? GetDirectChildValue(XElement parent, string localName)
        {
            return parent.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                ?.Value?.Trim();
        }

        private static string? GetFaultDetail(XElement fault)
        {
            var detail = fault.Elements().FirstOrDefault(e =>
                e.Name.LocalName.Equals("detail", StringComparison.OrdinalIgnoreCase));
            if (detail == null)
            {
                return null;
            }

            if (!detail.HasElements)
            {
                return string.IsNullOrWhiteSpace(detail.Value) ? null : detail.Value.Trim();
            }

            var bits = detail.Descendants()
                .Where(e => !e.HasElements && !string.IsNullOrWhiteSpace(e.Value))
                .Select(e => $"{e.Name.LocalName}: {e.Value.Trim()}")
                .ToList();

            return bits.Count == 0 ? null : string.Join("; ", bits);
        }

        private static string? ExtractErrorCode(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var match = Regex.Match(text, @"err_[a-z0-9_]+", RegexOptions.IgnoreCase);
            return match.Success ? match.Value.ToLowerInvariant() : null;
        }

        private static bool IsMissingConsignmentFault(string fault, XDocument document)
        {
            static bool Matches(string? text)
            {
                var code = ExtractErrorCode(text);
                return code is "err_cons_not_found" or "err_invalid_identifier";
            }

            return Matches(fault) || Matches(document.ToString());
        }

        private static bool IsGenericSoapCode(string value)
        {
            var normalized = value.Trim();
            if (normalized.Contains(':'))
            {
                normalized = normalized[(normalized.LastIndexOf(':') + 1)..];
            }

            return normalized.Equals("Client", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("Server", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("Sender", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("Receiver", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("VersionMismatch", StringComparison.OrdinalIgnoreCase)
                || normalized.Equals("MustUnderstand", StringComparison.OrdinalIgnoreCase);
        }

        private static string TranslateFault(string fault)
        {
            return fault switch
            {
                "err_user_incorrect_username_password"
                    => "Niepoprawna nazwa użytkownika i/lub hasło (err_user_incorrect_username_password).",
                "err_user_permissions_problem"
                    => "Problem z uprawnieniami użytkownika. Wymagany kontakt z GLS (err_user_permissions_problem).",
                "err_user_blocked"
                    => "Konto użytkownika w systemie GLS jest zablokowane (err_user_blocked).",
                "err_user_webapi_blocked"
                    => "Użytkownik nie posiada dostępu do usług WebAPI (err_user_webapi_blocked).",
                "err_user_login_by_localization_code_required"
                    => "Wymagane logowanie metodą adeLoginByLocalizationCode (err_user_login_by_localization_code_required).",
                "err_sess_create_problem"
                    => "Problem z procesem autoryzacji. Wymagany kontakt z GLS (err_sess_create_problem).",
                "err_sess_not_found"
                    => "Identyfikator sesji nie został znaleziony (err_sess_not_found).",
                "err_sess_expired"
                    => "Ważność identyfikatora sesji wygasła (err_sess_expired).",
                "err_user_insufficient_permissions"
                    => "Użytkownik nie posiada uprawnień do wykonania tej metody (err_user_insufficient_permissions).",
                "err_login_failed"
                    => "Niepoprawna nazwa użytkownika i/lub hasło (err_login_failed).",
                "err_cons_receiver_name1_empty"
                    => "Nazwa odbiorcy jest pusta (err_cons_receiver_name1_empty).",
                "err_cons_receiver_country_empty"
                    => "Kod kraju odbiorcy jest pusty (err_cons_receiver_country_empty).",
                "err_cons_receiver_zipcode_empty"
                    => "Kod pocztowy odbiorcy jest pusty (err_cons_receiver_zipcode_empty).",
                "err_cons_receiver_zipcode_is_invalid"
                    => "Kod pocztowy odbiorcy jest niepoprawny (err_cons_receiver_zipcode_is_invalid).",
                "err_cons_receiver_city_empty"
                    => "Miejscowość odbiorcy jest pusta (err_cons_receiver_city_empty).",
                "err_cons_receiver_street_empty"
                    => "Ulica odbiorcy jest pusta (err_cons_receiver_street_empty).",
                "err_cons_not_found"
                    => "Nie znaleziono przesyłki w przygotowalni GLS (err_cons_not_found).",
                "err_invalid_identifier"
                    => "Niepoprawny identyfikator przesyłki GLS (err_invalid_identifier).",
                "err_cons_no_label"
                    => "Przesyłka nie posiada etykiet (err_cons_no_label).",
                "err_cons_parcel_label_problem"
                    => "Problem z wygenerowaniem etykiety GLS (err_cons_parcel_label_problem).",
                "err_cons_currently_supported"
                    => "Przesyłka jest już obsługiwana (np. w doręczeniu) i nie można jej zmienić (err_cons_currently_supported).",
                "err_cons_delete_problem"
                    => "Wystąpił błąd usuwania przesyłki GLS (err_cons_delete_problem).",
                "err_cons_number_of_parcels_is_incorrect"
                    => "Liczba paczek w przesyłce jest nieprawidłowa (err_cons_number_of_parcels_is_incorrect).",
                "err_cons_srv_amount_of_cod_is_incorrect"
                    => "Kwota pobrania (COD) jest nieprawidłowa (err_cons_srv_amount_of_cod_is_incorrect).",
                "err_cons_srv_amount_of_cod_is_too_large"
                    => "Kwota pobrania (COD) jest za duża (err_cons_srv_amount_of_cod_is_too_large).",
                "err_cons_srv_cod_service_is_unavailable"
                    => "Usługa COD (za pobraniem) nie jest dostępna na tym koncie GLS (err_cons_srv_cod_service_is_unavailable).",
                "err_cons_srv_for_countries_other_than_pl_are_not_available"
                    => "Usługi dodatkowe (m.in. COD) są dostępne tylko dla odbiorców w Polsce (err_cons_srv_for_countries_other_than_pl_are_not_available).",
                "err_pickup_make_problem"
                    => "Problem z tworzeniem potwierdzenia nadania GLS (err_pickup_make_problem).",
                "err_pickup_empty_list_of_consignments"
                    => "Pusta lista przesyłek do potwierdzenia nadania (err_pickup_empty_list_of_consignments).",
                "err_pickup_one_consignment_id_is_invalid"
                    => "Niepoprawny identyfikator przesyłki — przesyłka mogła już trafić do potwierdzenia nadania (err_pickup_one_consignment_id_is_invalid).",
                "err_pickup_not_found"
                    => "Nie znaleziono potwierdzenia nadania GLS (err_pickup_not_found).",
                "err_id_start_invalid"
                    => "Niepoprawny identyfikator startowy GLS (err_id_start_invalid).",
                "err_user_debt_collection_lock"
                    => "Konto GLS ma blokadę windykacyjną — nie można dodawać paczek (err_user_debt_collection_lock).",
                _ => $"Błąd GLS: {fault}."
            };
        }

        private static string MaskSecrets(string xml)
        {
            var masked = Regex.Replace(
                xml,
                @"(<[^>]*user_password[^>]*>)(.*?)(</[^>]*user_password[^>]*>)",
                "$1***$3",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            return Regex.Replace(
                masked,
                @"(<[^>]*session[^>]*>)(.*?)(</[^>]*session[^>]*>)",
                "$1***$3",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }
    }
}
