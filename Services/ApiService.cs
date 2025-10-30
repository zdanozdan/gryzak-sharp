using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;

namespace Gryzak.Services
{
    public class ApiService
    {
        private readonly ConfigService _configService;
        private HttpClient? _httpClient;

        public ApiService(ConfigService configService)
        {
            _configService = configService;
        }

        public async Task<List<Order>> LoadOrdersAsync(int page, CancellationToken cancellationToken = default)
        {
            var config = _configService.LoadConfig();
            
            if (string.IsNullOrWhiteSpace(config.ApiUrl))
            {
                // Zwróć przykładowe dane testowe
                return GetTestData();
            }

            try
            {
                if (_httpClient == null)
                {
                    _httpClient = CreateHttpClient(config);
                }

                var url = BuildListUrl(config, page);
                var response = await _httpClient.GetAsync(url, cancellationToken);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync(cancellationToken);
                    var orders = ParseOrdersResponse(json);
                    return orders;
                }
                else
                {
                    throw new Exception($"API Error: {(int)response.StatusCode} {response.ReasonPhrase}");
                }
            }
            catch (TaskCanceledException)
            {
                throw new Exception("Timeout - połączenie przekroczyło limit czasu");
            }
            catch (HttpRequestException ex)
            {
                throw new Exception($"Błąd sieci - sprawdź konfigurację API: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Błąd ładowania zamówień: {ex.Message}");
            }
        }

        public async Task<string> GetOrderDetailsAsync(string orderId, CancellationToken cancellationToken = default)
        {
            var config = _configService.LoadConfig();
            
            if (string.IsNullOrWhiteSpace(config.ApiUrl))
            {
                // Zwróć przykładowe dane dla testów
                return $"{{ \"order_id\": \"{orderId}\", \"test\": true }}";
            }

            try
            {
                if (_httpClient == null)
                {
                    _httpClient = CreateHttpClient(config);
                }

                var url = BuildDetailsUrl(config, orderId);
                Console.WriteLine($"[ApiService] Wywoływanie URL: {url}");
                var response = await _httpClient.GetAsync(url, cancellationToken);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync(cancellationToken);
                    return json;
                }
                else
                {
                    throw new Exception($"API Error: {(int)response.StatusCode} {response.ReasonPhrase}");
                }
            }
            catch (TaskCanceledException)
            {
                throw new Exception("Timeout - połączenie przekroczyło limit czasu");
            }
            catch (HttpRequestException ex)
            {
                throw new Exception($"Błąd sieci - sprawdź konfigurację API: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Błąd ładowania szczegółów zamówienia: {ex.Message}");
            }
        }

        public async Task<bool> TestConnectionAsync(CancellationToken cancellationToken = default)
        {
            var config = _configService.LoadConfig();
            
            if (string.IsNullOrWhiteSpace(config.ApiUrl))
            {
                throw new Exception("URL API jest wymagany");
            }

            try
            {
                var client = CreateHttpClient(config);
                var url = BuildListUrl(config, 1);
                var response = await client.GetAsync(url, cancellationToken);
                
                return response.IsSuccessStatusCode;
            }
            catch (TaskCanceledException)
            {
                throw new Exception("Timeout - połączenie przekroczyło limit czasu");
            }
            catch (HttpRequestException)
            {
                throw new Exception("Błąd sieci - sprawdź URL i połączenie internetowe");
            }
            catch (Exception ex)
            {
                throw new Exception($"Błąd: {ex.Message}");
            }
        }

        private HttpClient CreateHttpClient(ApiConfig config)
        {
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(config.ApiTimeout);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            
            if (!string.IsNullOrWhiteSpace(config.ApiToken))
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config.ApiToken);
            }

            return client;
        }

        private string BuildListUrl(ApiConfig config, int page)
        {
            var baseUrl = config.ApiUrl.TrimEnd('/');
            var endpoint = config.OrderListEndpoint;
            var separator = endpoint.Contains('?') ? "&" : "?";
            return $"{baseUrl}{endpoint}{separator}page={page}";
        }

        private string BuildDetailsUrl(ApiConfig config, string orderId)
        {
            var baseUrl = config.ApiUrl.TrimEnd('/');
            var endpoint = config.OrderDetailsEndpoint.Replace("{order_id}", orderId);
            return $"{baseUrl}{endpoint}";
        }

        private List<Order> ParseOrdersResponse(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                List<JsonElement> ordersArray;

                // Format 1: Tablica bezpośrednio
                if (root.ValueKind == JsonValueKind.Array)
                {
                    ordersArray = root.EnumerateArray().ToList();
                }
                // Format 2: Obiekt z właściwością orders
                else if (root.TryGetProperty("orders", out var ordersProp))
                {
                    ordersArray = ordersProp.EnumerateArray().ToList();
                }
                // Format 3: Obiekt z właściwością data
                else if (root.TryGetProperty("data", out var dataProp))
                {
                    ordersArray = dataProp.EnumerateArray().ToList();
                }
                else
                {
                    return new List<Order>();
                }

                var result = new List<Order>();
                foreach (var orderElement in ordersArray)
                {
                    var order = MapToOrder(orderElement);
                    result.Add(order);
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Błąd parsowania odpowiedzi: {ex.Message}");
                return new List<Order>();
            }
        }

        private Order MapToOrder(JsonElement orderElement)
        {
            var orderId = orderElement.TryGetProperty("order_id", out var id) ? id.GetString() ?? "" : "";
            var firstname = orderElement.TryGetProperty("firstname", out var fn) ? fn.GetString() ?? "" : "";
            var lastname = orderElement.TryGetProperty("lastname", out var ln) ? ln.GetString() ?? "" : "";
            var email = orderElement.TryGetProperty("email", out var e) ? e.GetString() ?? "" : "";
            var telephone = orderElement.TryGetProperty("telephone", out var tel) ? tel.GetString() ?? "" : "";
            var company = orderElement.TryGetProperty("payment_company", out var comp) ? comp.GetString() : null;
            var vat = orderElement.TryGetProperty("vat", out var v) ? v.GetString() : null;
            var address1 = orderElement.TryGetProperty("payment_address_1", out var a1) ? a1.GetString() ?? "" : "";
            var address2 = orderElement.TryGetProperty("payment_address_2", out var a2) ? a2.GetString() : null;
            var postcode = orderElement.TryGetProperty("payment_postcode", out var pc) ? pc.GetString() ?? "" : "";
            var city = orderElement.TryGetProperty("payment_city", out var c) ? c.GetString() ?? "" : "";
            var status = orderElement.TryGetProperty("status", out var s) ? s.GetString() ?? "" : "";
            var paymentStatus = orderElement.TryGetProperty("payment_status", out var ps) ? ps.GetString() ?? "Nieznany" : "Nieznany";
            
            // Obsługa total jako liczby float/double lub stringa
            string total = "0";
            if (orderElement.TryGetProperty("total", out var t))
            {
                if (t.ValueKind == JsonValueKind.Number)
                {
                    var totalValue = t.GetDouble();
                    total = totalValue.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                    Console.WriteLine($"[ApiService] Odczytano total jako liczbę: {totalValue} -> {total}");
                }
                else if (t.ValueKind == JsonValueKind.String)
                {
                    total = t.GetString() ?? "0";
                    Console.WriteLine($"[ApiService] Odczytano total jako string: {total}");
                }
                else
                {
                    Console.WriteLine($"[ApiService] Total ma nieoczekiwany typ: {t.ValueKind}");
                }
            }
            else
            {
                Console.WriteLine("[ApiService] Pole 'total' nie znalezione w odpowiedzi API");
            }
            
            var currency = orderElement.TryGetProperty("currency_code", out var curr) ? curr.GetString() ?? "PLN" : "PLN";
            var dateAdded = orderElement.TryGetProperty("date_added", out var date) ? date.GetString() ?? "" : "";

            var address = "";
            if (!string.IsNullOrWhiteSpace(address1))
            {
                address = address1;
                if (!string.IsNullOrWhiteSpace(address2))
                {
                    address += ", " + address2;
                }
                address += $", {postcode} {city}";
            }

            DateTime dateParsed = DateTime.Now;
            if (!string.IsNullOrWhiteSpace(dateAdded))
            {
                if (DateTime.TryParse(dateAdded, out var parsed))
                {
                    dateParsed = parsed;
                }
            }

            // Sformatuj total - parsuj i formatuj z InvariantCulture
            string formattedTotal = "0.00";
            if (!string.IsNullOrWhiteSpace(total) && total != "0")
            {
                if (double.TryParse(total, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out var totalVal))
                {
                    formattedTotal = totalVal.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
                    Console.WriteLine($"[ApiService] Zamówienie {orderId}: total '{total}' -> sformatowane '{formattedTotal}'");
                }
                else
                {
                    Console.WriteLine($"[ApiService] BŁĄD: Nie można sparsować total '{total}' dla zamówienia {orderId}");
                    formattedTotal = total; // Użyj oryginalnej wartości jako fallback
                }
            }
            else
            {
                Console.WriteLine($"[ApiService] Zamówienie {orderId}: total jest pusty lub 0");
            }

            // Dekoduj encje HTML (czasem podwójnie zakodowane)
            if (!string.IsNullOrWhiteSpace(company))
            {
                company = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(company));
            }
            address1 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address1));
            if (!string.IsNullOrWhiteSpace(address2)) address2 = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(address2));
            postcode = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(postcode));
            city = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(city));

            return new Order
            {
                Id = orderId,
                Customer = $"{firstname} {lastname}".Trim(),
                Email = string.IsNullOrWhiteSpace(email) ? "Brak email" : email,
                Phone = string.IsNullOrWhiteSpace(telephone) ? "Brak telefonu" : telephone,
                Company = company,
                Nip = vat,
                Address = string.IsNullOrWhiteSpace(address) ? null : address,
                Status = status,
                PaymentStatus = paymentStatus,
                Total = formattedTotal,
                Currency = currency,
                Date = dateParsed
            };
        }

        private List<Order> GetTestData()
        {
            return new List<Order>
            {
                new Order
                {
                    Id = "12345",
                    Customer = "Jan Kowalski",
                    Email = "jan.kowalski@example.com",
                    Phone = "+48123456789",
                    Company = "Firma Sp. z o.o.",
                    Nip = "PL1234567890",
                    Address = "ul. Przykładowa 123, lok. 45, 00-001 Warszawa",
                    Status = "Nowe",
                    PaymentStatus = "Zapłacono",
                    Total = "123.45",
                    Currency = "PLN",
                    Date = DateTime.Parse("2024-01-15T10:30:00Z")
                },
                new Order
                {
                    Id = "12346",
                    Customer = "Anna Nowak",
                    Email = "anna.nowak@example.com",
                    Phone = "+48987654321",
                    Company = null,
                    Nip = null,
                    Address = "ul. Testowa 456, 02-002 Kraków",
                    Status = "Wysłano",
                    PaymentStatus = "Zapłacono",
                    Total = "456.78",
                    Currency = "PLN",
                    Date = DateTime.Parse("2024-01-16T14:20:00Z")
                },
                new Order
                {
                    Id = "12347",
                    Customer = "Piotr Wiśniewski",
                    Email = "piotr.wisniewski@example.com",
                    Phone = "+48555123456",
                    Company = "Inna Firma Ltd",
                    Nip = "PL9876543210",
                    Address = "ul. Przykładowa 789, 03-003 Gdańsk",
                    Status = "Dostarczone",
                    PaymentStatus = "Zapłacono",
                    Total = "789.12",
                    Currency = "PLN",
                    Date = DateTime.Parse("2024-01-17T09:15:00Z")
                }
            };
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}

