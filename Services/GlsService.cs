using System;
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

        private async Task<XDocument> SendSoapRequestAsync(
            string operation,
            CancellationToken cancellationToken,
            params (string Name, string Value)[] parameters)
        {
            var apiUrl = _config.GetActiveApiUrl();
            var soapXml = BuildSoapEnvelope(apiUrl, operation, parameters);
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
        private static string BuildSoapEnvelope(string apiUrl, string operation, (string Name, string Value)[] parameters)
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
            foreach (var (name, value) in parameters)
            {
                sb.Append('<').Append(name).Append(" xsi:type=\"xsd:string\">");
                sb.Append(SecurityElement.Escape(value) ?? "");
                sb.Append("</").Append(name).Append('>');
            }
            sb.Append("</ns1:").Append(operation).Append('>');
            sb.Append("</SOAP-ENV:Body>");
            sb.Append("</SOAP-ENV:Envelope>");
            return sb.ToString();
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
                    ? "Błąd logowania GLS."
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
