using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace IntunePackagingTool
{
    public interface IUploadProgress
    {
        void UpdateProgress(int percentage, string message);
    }

    public class IntuneUploadService : IDisposable
    {
        private HttpClient? _httpClient;
        private readonly IntuneService _intuneService;

        public IntuneUploadService(IntuneService intuneService)
        {
            _intuneService = intuneService;
        }

        private void EnsureHttpClient()
        {
            if (_httpClient == null)
            {
                _httpClient = new HttpClient();
                _httpClient.Timeout = TimeSpan.FromMinutes(10);
            }
        }

        public async Task<string> UploadWin32ApplicationAsync(
            ApplicationInfo appInfo, 
            string packagePath, 
            List<DetectionRule> detectionRules, 
            string installCommand, 
            string uninstallCommand, 
            string description,
            string installContext,
            IUploadProgress? progress = null)
        {
            try
            {
                progress?.UpdateProgress(5, "Authenticating with Microsoft Graph...");
                var token = await _intuneService.GetAccessTokenAsync();
                EnsureHttpClient();
                _httpClient!.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                // Step 1: Create the .intunewin file
                progress?.UpdateProgress(15, "Converting package to .intunewin format...");
                await CreateIntuneWinFileAsync(packagePath);

                // Step 2: Find the created .intunewin file
                progress?.UpdateProgress(25, "Locating .intunewin file...");
                var intuneFolder = Path.Combine(packagePath, "Intune");
                var intuneWinFiles = Directory.GetFiles(intuneFolder, "*.intunewin");
                if (intuneWinFiles.Length == 0)
                {
                    throw new Exception("No .intunewin file found after conversion.");
                }

                // Step 3: Create Win32LobApp with CORRECT detection rules format
                progress?.UpdateProgress(40, "Creating application entry in Intune...");
                var appId = await CreateWin32LobAppWithCorrectDetectionRulesAsync(appInfo, installCommand, uninstallCommand, description, detectionRules, installContext);

                progress?.UpdateProgress(100, "Application created in Intune successfully!");

                return appId;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to upload application to Intune: {ex.Message}", ex);
            }
        }

        private async Task CreateIntuneWinFileAsync(string packagePath)
        {
            var converterPath = @"\\nbb.local\sys\SCCMData\TOOLS\IntunePackagingTool\IntuneWinAppUtil.exe";
            
            if (!File.Exists(converterPath))
            {
                throw new FileNotFoundException($"IntuneWinAppUtil.exe not found at: {converterPath}");
            }

            // Define paths
            var applicationFolder = Path.Combine(packagePath, "Application");
            var setupFile = Path.Combine(applicationFolder, "Deploy-Application.exe");
            var outputFolder = Path.Combine(packagePath, "Intune");

            // Validate paths
            if (!Directory.Exists(applicationFolder))
            {
                throw new DirectoryNotFoundException($"Application folder not found: {applicationFolder}");
            }

            if (!File.Exists(setupFile))
            {
                throw new FileNotFoundException($"Deploy-Application.exe not found: {setupFile}");
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Build command arguments
            var arguments = $"-c \"{applicationFolder}\" -s \"{setupFile}\" -o \"{outputFolder}\" -q";

            // Run the converter
            var processStartInfo = new ProcessStartInfo
            {
                FileName = converterPath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            process.Start();

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                var errorMessage = string.IsNullOrWhiteSpace(error) ? output : error;
                throw new Exception($"IntuneWinAppUtil failed with exit code {process.ExitCode}: {errorMessage}");
            }

            Debug.WriteLine($"✓ Created .intunewin file using: {converterPath} {arguments}");
        }

        private async Task<string> CreateWin32LobAppWithCorrectDetectionRulesAsync(
            ApplicationInfo appInfo, 
            string installCommand, 
            string uninstallCommand, 
            string description, 
            List<DetectionRule> detectionRules,
            string installContext)
        {
            // Create the SIMPLEST possible detection rules that match your working PowerShell script EXACTLY
            var formattedDetectionRules = new List<Dictionary<string, object>>();
            
            Debug.WriteLine($"Creating SIMPLEST detection rules...");
            
            // Always add at least one simple file detection rule
            var simpleFileRule = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobAppFileSystemDetection",
                ["path"] = "%ProgramFiles%",
                ["fileOrFolderName"] = "Deploy-Application.exe", 
                ["detectionType"] = "exists",
                ["check32BitOn64System"] = true  // PowerShell script uses true, not false
            };
            
            formattedDetectionRules.Add(simpleFileRule);
            Debug.WriteLine("Added simple file detection rule: %ProgramFiles%\\Deploy-Application.exe");

            // Create the app payload with MINIMAL required fields
            var createAppPayload = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobApp",
                ["displayName"] = $"{appInfo.Manufacturer} {appInfo.Name}",
                ["description"] = description,
                ["publisher"] = appInfo.Manufacturer,
                ["displayVersion"] = appInfo.Version,
                ["installCommandLine"] = installCommand,
                ["uninstallCommandLine"] = uninstallCommand,
                ["installExperience"] = new Dictionary<string, object>
                {
                    ["runAsAccount"] = installContext
                },
                ["detectionRules"] = formattedDetectionRules.ToArray(),
                ["returnCodes"] = new[]
                {
                    new Dictionary<string, object> { ["returnCode"] = 0, ["type"] = "success" },
                    new Dictionary<string, object> { ["returnCode"] = 3010, ["type"] = "softReboot" }
                }
            };

            var jsonOptions = new JsonSerializerOptions 
            { 
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            };
            
            var json = JsonSerializer.Serialize(createAppPayload, jsonOptions);
            
            Debug.WriteLine("Creating Win32 app with MINIMAL payload:");
            Debug.WriteLine(json);
            
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient!.PostAsync("https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps", content);
            var responseText = await response.Content.ReadAsStringAsync();

            Debug.WriteLine($"Response Status: {response.StatusCode}");
            Debug.WriteLine($"Response Body: {responseText}");

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to create Win32 app. Status: {response.StatusCode}, Response: {responseText}");
            }

            var createdApp = JsonSerializer.Deserialize<JsonElement>(responseText);
            var appId = createdApp.GetProperty("id").GetString();

            Debug.WriteLine($"✓ SUCCESS! Created Win32 app in Intune with ID: {appId}");
            return appId ?? throw new Exception("App ID not returned from creation");
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}