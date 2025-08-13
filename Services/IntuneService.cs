using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Diagnostics;
using System.Linq;

namespace IntunePackagingTool
{
    public class IntuneService
    {
        // Remove static to allow proper disposal and recreation
        private HttpClient? _httpClient;
        private string? _accessToken;
        private DateTime _tokenExpiry = DateTime.MinValue;
        private bool _enableDebug = false;

        private readonly string _clientId = "b47987a1-70b4-415a-9a4e-9775473e382b";
        private readonly string _tenantId = "43f10d24-b9bf-46da-a9c8-15c1b0990ce7";
        private readonly string _certificateThumbprint = "CF6DCE7DF3377CA65D9B40F06BF8C2228AC7821F";

        public void EnableDebug(bool enable)
        {
            _enableDebug = enable;
        }

        private void ShowDebug(string message, string title = "Debug")
        {
            if (_enableDebug)
            {
                MessageBox.Show(message, title);
            }
        }

        private void EnsureHttpClient()
        {
            if (_httpClient == null)
            {
                _httpClient = new HttpClient();
                _httpClient.Timeout = TimeSpan.FromMinutes(2);
            }
        }

        public async Task<string> GetAccessTokenAsync()
        {
            if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _tokenExpiry.AddMinutes(-5))
            {
                ShowDebug("Using cached access token", "Debug - Step 1");
                return _accessToken;
            }

            EnsureHttpClient();
            var certificate = LoadCertificate();
            var assertion = CreateJwtAssertion(certificate);

            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"),
                new KeyValuePair<string, string>("client_assertion", assertion),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                new KeyValuePair<string, string>("grant_type", "client_credentials")
            });

            var tokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
            var response = await _httpClient!.PostAsync(tokenUrl, content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                ShowDebug($"Token request failed: {response.StatusCode}\n{responseText}", "Debug - Token Error");
                throw new Exception($"Failed to get access token. Status: {response.StatusCode}, Response: {responseText}");
            }

            var tokenResponse = JsonSerializer.Deserialize<TokenResponse>(responseText);
            if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.AccessToken))
            {
                ShowDebug("Access token is missing in response", "Debug - Token Error");
                throw new Exception("Access token is missing in response");
            }

            _accessToken = tokenResponse.AccessToken;
            _tokenExpiry = DateTime.UtcNow.AddSeconds(tokenResponse.ExpiresIn);

            ShowDebug("Access token obtained successfully", "Debug - Success");
            return _accessToken;
        }

        private X509Certificate2 LoadCertificate()
        {
            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, _certificateThumbprint, false);
            if (certificates.Count == 0)
            {
                ShowDebug("Certificate not found", "Debug - Certificate Error");
                throw new Exception("Certificate not found");
            }
            return certificates[0];
        }

        private string CreateJwtAssertion(X509Certificate2 certificate)
        {
            var now = DateTimeOffset.UtcNow;
            var x5t = Base64UrlEncode(certificate.GetCertHash());

            var header = new
            {
                alg = "RS256",
                typ = "JWT",
                x5t = x5t
            };

            var payload = new
            {
                aud = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token",
                exp = now.AddMinutes(10).ToUnixTimeSeconds(),
                iss = _clientId,
                jti = Guid.NewGuid().ToString(),
                nbf = now.ToUnixTimeSeconds(),
                sub = _clientId
            };

            var headerJson = JsonSerializer.Serialize(header);
            var payloadJson = JsonSerializer.Serialize(payload);

            var headerEncoded = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
            var payloadEncoded = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));

            var message = $"{headerEncoded}.{payloadEncoded}";
            var messageBytes = Encoding.UTF8.GetBytes(message);

            using var rsa = certificate.GetRSAPrivateKey();
            var signature = rsa.SignData(messageBytes, System.Security.Cryptography.HashAlgorithmName.SHA256, System.Security.Cryptography.RSASignaturePadding.Pkcs1);
            var signatureEncoded = Base64UrlEncode(signature);

            return $"{message}.{signatureEncoded}";
        }

        private static string Base64UrlEncode(byte[] input)
        {
            return Convert.ToBase64String(input)
                .Replace("+", "-")
                .Replace("/", "_")
                .Replace("=", "");
        }

        public async Task<List<IntuneApplication>> GetApplicationsAsync()
        {
            try
            {
                Debug.WriteLine("=== GRAPH API CALL ===");
                Debug.WriteLine("Starting GetApplicationsAsync...");
                
                var token = await GetAccessTokenAsync();
                Debug.WriteLine($"✓ Token obtained for Graph API call, length: {token.Length}");
                
                EnsureHttpClient();
                _httpClient!.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                // Try with smaller batch and expand categories - if it fails, we'll fallback
                var requestUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$filter=isof('microsoft.graph.win32LobApp')&$expand=categories&$top=300";
                Debug.WriteLine($"Making Graph API request to: {requestUrl}");

                var response = await _httpClient.GetAsync(requestUrl);
                var responseText = await response.Content.ReadAsStringAsync();

                Debug.WriteLine($"Graph API response status: {response.StatusCode}");
                Debug.WriteLine($"Graph API response length: {responseText.Length}");

                if (!response.IsSuccessStatusCode)
                {
                    Debug.WriteLine($"✗ Graph API error with categories expand, trying without expand...");
                    
                    // Fallback: try without expand if categories expand fails
                    requestUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$filter=isof('microsoft.graph.win32LobApp')&$top=100";
                    Debug.WriteLine($"Fallback request to: {requestUrl}");
                    
                    response = await _httpClient.GetAsync(requestUrl);
                    responseText = await response.Content.ReadAsStringAsync();
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        Debug.WriteLine($"✗ Graph API error response: {responseText}");
                        
                        // Try to parse Graph API error
                        try
                        {
                            var errorResponse = JsonSerializer.Deserialize<JsonElement>(responseText);
                            if (errorResponse.TryGetProperty("error", out var errorObj))
                            {
                                var code = errorObj.TryGetProperty("code", out var codeProp) ? codeProp.GetString() : "Unknown";
                                var message = errorObj.TryGetProperty("message", out var msgProp) ? msgProp.GetString() : "Unknown";
                                
                                Debug.WriteLine($"Graph API Error - Code: {code}, Message: {message}");
                                throw new Exception($"Graph API Error: {code} - {message}");
                            }
                        }
                        catch (JsonException)
                        {
                            Debug.WriteLine("Could not parse Graph API error response");
                        }
                        
                        throw new Exception($"Graph API request failed. Status: {response.StatusCode}, Response: {responseText}");
                    }
                }

                Debug.WriteLine("✓ Graph API request successful, parsing response...");

                var graphResponse = JsonSerializer.Deserialize<GraphResponse<JsonElement>>(responseText, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
                var apps = new List<IntuneApplication>();

                if (graphResponse?.Value != null)
                {
                    Debug.WriteLine($"✓ Found {graphResponse.Value.Count} applications in Intune");
                    
                    foreach (var app in graphResponse.Value)
                    {
                        var id = app.GetProperty("id").GetString() ?? "";
                        var displayName = app.TryGetProperty("displayName", out var dnProp) ? dnProp.GetString() ?? "Unknown" : "Unknown";
                        var version = app.TryGetProperty("displayVersion", out var verProp) ? verProp.GetString() ?? "1.0.0" : "1.0.0";
                        var publisher = app.TryGetProperty("publisher", out var pubProp) ? pubProp.GetString() ?? "Unknown" : "Unknown";

                        // Try to get real categories from Intune if they were expanded
                        var category = "Business"; // Default
                        if (app.TryGetProperty("categories", out var categoriesProp) && categoriesProp.ValueKind == JsonValueKind.Array)
                        {
                            var categoryNames = new List<string>();
                            foreach (var cat in categoriesProp.EnumerateArray())
                            {
                                if (cat.TryGetProperty("displayName", out var catName))
                                {
                                    var catNameStr = catName.GetString();
                                    if (!string.IsNullOrWhiteSpace(catNameStr))
                                        categoryNames.Add(catNameStr);
                                }
                            }
                            if (categoryNames.Count > 0)
                                category = string.Join(", ", categoryNames);
                        }

                        Debug.WriteLine($"  App: {displayName} v{version} by {publisher} - Category: {category}");

                        apps.Add(new IntuneApplication
                        {
                            Id = id,
                            DisplayName = displayName,
                            Version = version,
                            Publisher = publisher,
                            Category = category,
                            LastModified = DateTime.Now
                        });
                    }
                }
                else
                {
                    Debug.WriteLine("⚠ No applications found in response");
                }

                // Handle pagination if there are more results (but limit to avoid timeouts)
                var hasNextPage = graphResponse?.ODataNextLink != null;
                var pageCount = 1;
                
                while (hasNextPage && pageCount < 20) // Allow more pages since we're using smaller batches
                {
                    Debug.WriteLine($"Getting page {pageCount + 1}...");
                    
                    var nextResponse = await _httpClient.GetAsync(graphResponse!.ODataNextLink);
                    if (!nextResponse.IsSuccessStatusCode) break;
                    
                    var nextResponseText = await nextResponse.Content.ReadAsStringAsync();
                    var nextGraphResponse = JsonSerializer.Deserialize<GraphResponse<JsonElement>>(nextResponseText, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
                    
                    if (nextGraphResponse?.Value != null)
                    {
                        Debug.WriteLine($"✓ Found {nextGraphResponse.Value.Count} more applications");
                        
                        foreach (var app in nextGraphResponse.Value)
                        {
                            var id = app.GetProperty("id").GetString() ?? "";
                            var displayName = app.TryGetProperty("displayName", out var dnProp) ? dnProp.GetString() ?? "Unknown" : "Unknown";
                            var version = app.TryGetProperty("displayVersion", out var verProp) ? verProp.GetString() ?? "1.0.0" : "1.0.0";
                            var publisher = app.TryGetProperty("publisher", out var pubProp) ? pubProp.GetString() ?? "Unknown" : "Unknown";

                            // Try to get real categories
                            var category = "Business"; // Default
                            if (app.TryGetProperty("categories", out var categoriesProp) && categoriesProp.ValueKind == JsonValueKind.Array)
                            {
                                var categoryNames = new List<string>();
                                foreach (var cat in categoriesProp.EnumerateArray())
                                {
                                    if (cat.TryGetProperty("displayName", out var catName))
                                    {
                                        var catNameStr = catName.GetString();
                                        if (!string.IsNullOrWhiteSpace(catNameStr))
                                            categoryNames.Add(catNameStr);
                                    }
                                }
                                if (categoryNames.Count > 0)
                                    category = string.Join(", ", categoryNames);
                            }

                            apps.Add(new IntuneApplication
                            {
                                Id = id,
                                DisplayName = displayName,
                                Version = version,
                                Publisher = publisher,
                                Category = category,
                                LastModified = DateTime.Now
                            });
                        }
                    }
                    
                    graphResponse = nextGraphResponse;
                    hasNextPage = graphResponse?.ODataNextLink != null;
                    pageCount++;
                }

                Debug.WriteLine($"✓ Successfully processed {apps.Count} applications across {pageCount} pages");
                return apps.OrderBy(a => a.DisplayName).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"✗ Error in GetApplicationsAsync: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                
                MessageBox.Show(
                    $"Failed to retrieve applications from Intune:\n\n{ex.Message}\n\nCheck the Debug Output window for detailed logs.",
                    "Graph API Error", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
                
                throw;
            }
        }

        public async Task RunFullDebugTestAsync()
        {
            EnableDebug(true);

            try
            {
                ShowDebug("Starting comprehensive debug test...", "Debug Test");

                var token = await GetAccessTokenAsync();
                ShowDebug($"✓ Token obtained. Length: {token.Length}", "Debug Test");

                var apps = await GetApplicationsAsync();
                ShowDebug($"✓ Retrieved {apps.Count} Win32 applications from Intune.", "Debug Test");

                ShowDebug("🎉 COMPLETE SUCCESS! All tests passed.", "Debug Test");
            }
            catch (Exception ex)
            {
                ShowDebug($"❌ Debug test failed: {ex.Message}", "Debug Test");
            }
        }

        public async Task UploadApplicationAsync(ApplicationInfo appInfo)
        {
            try
            {
                var token = await GetAccessTokenAsync();
                EnsureHttpClient();
                _httpClient!.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                await Task.Delay(1000); // Simulate upload delay

                MessageBox.Show(
                    $"Win32 app upload via API requires complex file handling.\n\n" +
                    $"Package: {appInfo.Manufacturer}_{appInfo.Name}_{appInfo.Version}\n\n" +
                    $"Package is ready at the network location.\n" +
                    $"Use Intune admin center for upload, or we can implement the upload API later.",
                    "Upload Status", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to upload application to Intune: {ex.Message}", ex);
            }
        }

        // Dispose method to clean up HttpClient
        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    public class GraphResponse<T>
    {
        [JsonPropertyName("@odata.context")]
        public string? ODataContext { get; set; }

        [JsonPropertyName("@odata.nextLink")]
        public string? ODataNextLink { get; set; }

        [JsonPropertyName("value")]
        public List<T>? Value { get; set; }
    }

    public class TokenResponse
    {
        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; } = "";

        [JsonPropertyName("token_type")]
        public string TokenType { get; set; } = "";

        [JsonPropertyName("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonPropertyName("scope")]
        public string Scope { get; set; } = "";
    }

}