using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;

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
                _httpClient.Timeout = TimeSpan.FromMinutes(30); // Increased for file uploads
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
                progress?.UpdateProgress(10, "Converting package to .intunewin format...");
                await CreateIntuneWinFileAsync(packagePath);

                // Step 2: Find the created .intunewin file
                progress?.UpdateProgress(15, "Locating .intunewin file...");
                var intuneFolder = Path.Combine(packagePath, "Intune");
                var intuneWinFiles = Directory.GetFiles(intuneFolder, "*.intunewin");
                if (intuneWinFiles.Length == 0)
                {
                    throw new Exception("No .intunewin file found after conversion.");
                }
                
                var intuneWinFile = intuneWinFiles[0];
                Debug.WriteLine($"✓ Found .intunewin file: {Path.GetFileName(intuneWinFile)}");

                // Step 3: Extract .intunewin metadata
                progress?.UpdateProgress(20, "Extracting .intunewin metadata...");
                var intuneWinInfo = ExtractIntuneWinInfo(intuneWinFile);

                // Step 4: Create Win32LobApp
                progress?.UpdateProgress(25, "Creating application in Intune...");
                var appId = await CreateWin32LobAppAsync(appInfo, installCommand, uninstallCommand, description, detectionRules, installContext, intuneWinInfo);

                // Step 5: Create content version
                progress?.UpdateProgress(35, "Creating content version...");
                var contentVersionId = await CreateContentVersionAsync(appId);

                // Step 6: Create file entry
                progress?.UpdateProgress(45, "Creating file entry...");
                var fileId = await CreateFileEntryAsync(appId, contentVersionId, intuneWinInfo);

                // Step 7: Wait for Azure Storage URI
                progress?.UpdateProgress(55, "Getting Azure Storage URI...");
                var azureStorageInfo = await WaitForAzureStorageUriAsync(appId, contentVersionId, fileId);

                // Step 8: Upload file to Azure Storage
                progress?.UpdateProgress(65, "Uploading file to Azure Storage...");
                await UploadFileToAzureStorageAsync(azureStorageInfo.SasUri, intuneWinInfo.EncryptedFilePath, progress);

                // Step 9: Commit the file
                progress?.UpdateProgress(85, "Committing file...");
                await CommitFileAsync(appId, contentVersionId, fileId, intuneWinInfo.EncryptionInfo);

                // Step 10: Wait for file processing
                progress?.UpdateProgress(90, "Waiting for file processing...");
                await WaitForFileProcessingAsync(appId, contentVersionId, fileId, "CommitFile");

                // Step 11: Commit the app
                progress?.UpdateProgress(95, "Finalizing application...");
                await CommitAppAsync(appId, contentVersionId);

                // Step 12: Cleanup temp files
                CleanupTempFiles(intuneWinInfo);

                progress?.UpdateProgress(100, "Application uploaded successfully!");
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

            var applicationFolder = Path.Combine(packagePath, "Application");
            var setupFile = Path.Combine(applicationFolder, "Deploy-Application.exe");
            var outputFolder = Path.Combine(packagePath, "Intune");

            if (!Directory.Exists(applicationFolder))
            {
                throw new DirectoryNotFoundException($"Application folder not found: {applicationFolder}");
            }

            if (!File.Exists(setupFile))
            {
                throw new FileNotFoundException($"Deploy-Application.exe not found: {setupFile}");
            }

            Directory.CreateDirectory(outputFolder);

            var arguments = $"-c \"{applicationFolder}\" -s \"{setupFile}\" -o \"{outputFolder}\" -q";

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

      private IntuneWinInfo ExtractIntuneWinInfo(string intuneWinFilePath)
{
    var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
    Directory.CreateDirectory(tempDir);

    try
    {
        using (var archive = ZipFile.OpenRead(intuneWinFilePath))
        {
            // Find detection.xml (case insensitive)
            var detectionEntry = archive.Entries.FirstOrDefault(e => 
                e.Name.Equals("detection.xml", StringComparison.OrdinalIgnoreCase));
            
            if (detectionEntry == null)
            {
                var availableFiles = string.Join(", ", archive.Entries.Select(e => e.Name));
                throw new Exception($"detection.xml not found in .intunewin file. Available files: {availableFiles}");
            }

            // Extract detection.xml
            var detectionXmlPath = Path.Combine(tempDir, "detection.xml");
            detectionEntry.ExtractToFile(detectionXmlPath);

            // Read and parse the XML (handling namespaces)
            var xmlContent = File.ReadAllText(detectionXmlPath);
            Debug.WriteLine("=== DETECTION.XML CONTENT ===");
            Debug.WriteLine(xmlContent);

            var detectionXml = XDocument.Load(detectionXmlPath);
            
            // Handle namespaces properly - get the root element regardless of namespace
            var appInfo = detectionXml.Root;
            if (appInfo == null || !appInfo.Name.LocalName.Equals("ApplicationInfo", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception($"Root element is not ApplicationInfo. Found: {appInfo?.Name}");
            }

            Debug.WriteLine("✓ Found ApplicationInfo root element");

            // Get elements by LocalName to ignore namespaces
            var encryptionInfo = appInfo.Elements().FirstOrDefault(e => e.Name.LocalName == "EncryptionInfo");
            if (encryptionInfo == null)
            {
                var availableElements = string.Join(", ", appInfo.Elements().Select(e => e.Name.LocalName));
                throw new Exception($"EncryptionInfo not found. Available elements: {availableElements}");
            }

            Debug.WriteLine("✓ Found EncryptionInfo element");

            // Extract the encrypted content file (try both names)
            var contentsEntry = archive.Entries.FirstOrDefault(e => 
                e.Name.Equals("Contents.dat", StringComparison.OrdinalIgnoreCase) ||
                e.Name.Equals("IntunePackage.dat", StringComparison.OrdinalIgnoreCase));
            
            if (contentsEntry == null)
            {
                // Look for any .dat file
                contentsEntry = archive.Entries.FirstOrDefault(e => e.Name.EndsWith(".dat", StringComparison.OrdinalIgnoreCase));
                
                if (contentsEntry == null)
                {
                    throw new Exception("No encrypted content file (.dat) found in .intunewin archive");
                }
            }

            var encryptedFilePath = Path.Combine(tempDir, contentsEntry.Name);
            contentsEntry.ExtractToFile(encryptedFilePath);
            Debug.WriteLine($"✓ Extracted encrypted content: {contentsEntry.Name}");

            // Helper function to get element value by LocalName (ignoring namespaces)
            string GetElementValue(XElement parent, string localName)
            {
                var element = parent.Elements().FirstOrDefault(e => e.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase));
                var value = element?.Value?.Trim();
                
                if (string.IsNullOrWhiteSpace(value))
                {
                    Debug.WriteLine($"⚠ Element '{localName}' is empty or missing");
                    return "";
                }
                
                Debug.WriteLine($"✓ Element '{localName}': {value}");
                return value;
            }

            // Extract values using LocalName to ignore namespaces
            var fileName = GetElementValue(appInfo, "FileName");
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = Path.GetFileName(intuneWinFilePath);
                Debug.WriteLine($"Using fallback filename: {fileName}");
            }

            var unencryptedSizeStr = GetElementValue(appInfo, "UnencryptedContentSize");
            if (!long.TryParse(unencryptedSizeStr, out var unencryptedSize))
            {
                Debug.WriteLine($"⚠ Could not parse UnencryptedContentSize: '{unencryptedSizeStr}', using encrypted file size as fallback");
                unencryptedSize = new FileInfo(encryptedFilePath).Length;
            }

            var result = new IntuneWinInfo
            {
                FileName = fileName,
                UnencryptedContentSize = unencryptedSize,
                EncryptedFilePath = encryptedFilePath,
                TempDirectory = tempDir,
                EncryptionInfo = new EncryptionInfo
                {
                    EncryptionKey = GetElementValue(encryptionInfo, "EncryptionKey"),
                    MacKey = GetElementValue(encryptionInfo, "MacKey") ?? GetElementValue(encryptionInfo, "macKey"),
                    InitializationVector = GetElementValue(encryptionInfo, "InitializationVector") ?? GetElementValue(encryptionInfo, "initializationVector"),
                    Mac = GetElementValue(encryptionInfo, "Mac") ?? GetElementValue(encryptionInfo, "mac"),
                    ProfileIdentifier = "ProfileVersion1",
                    FileDigest = GetElementValue(encryptionInfo, "FileDigest") ?? GetElementValue(encryptionInfo, "fileDigest"),
                    FileDigestAlgorithm = GetElementValue(encryptionInfo, "FileDigestAlgorithm") ?? GetElementValue(encryptionInfo, "fileDigestAlgorithm") ?? "SHA256"
                }
            };

            // Validate that we got the essential encryption info
            if (string.IsNullOrEmpty(result.EncryptionInfo.EncryptionKey))
            {
                throw new Exception("EncryptionKey is missing from detection.xml");
            }

            Debug.WriteLine("✓ Successfully extracted IntuneWin metadata");
            Debug.WriteLine($"  FileName: {result.FileName}");
            Debug.WriteLine($"  UnencryptedSize: {result.UnencryptedContentSize:N0} bytes");
            Debug.WriteLine($"  EncryptionKey: {result.EncryptionInfo.EncryptionKey[..Math.Min(10, result.EncryptionInfo.EncryptionKey.Length)]}...");
            
            return result;
        }
    }
    catch
    {
        // Clean up temp directory if extraction fails
        if (Directory.Exists(tempDir))
        {
            try
            {
                Directory.Delete(tempDir, true);
            }
            catch { }
        }
        throw;
    }
}

        private async Task<string> CreateWin32LobAppAsync(
            ApplicationInfo appInfo, 
            string installCommand, 
            string uninstallCommand, 
            string description, 
            List<DetectionRule> detectionRules,
            string installContext,
            IntuneWinInfo intuneWinInfo)
        {
            var formattedDetectionRules = new List<Dictionary<string, object>>();
            
            foreach (var rule in detectionRules)
            {
                var formattedRule = ConvertDetectionRuleForBetaAPI(rule);
                if (formattedRule != null)
                {
                    formattedDetectionRules.Add(formattedRule);
                }
            }

            if (formattedDetectionRules.Count == 0)
            {
                // Add default detection rule if none provided
                formattedDetectionRules.Add(new Dictionary<string, object>
                {
                    ["@odata.type"] = "#microsoft.graph.win32LobAppFileSystemDetection",
                    ["path"] = "%ProgramFiles%",
                    ["fileOrFolderName"] = $"{appInfo.Name}.exe",
                    ["check32BitOn64System"] = false,
                    ["detectionType"] = "exists"
                });
            }

            var createAppPayload = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobApp",
                ["displayName"] = $"{appInfo.Manufacturer} {appInfo.Name}",
                ["description"] = description,
                ["publisher"] = appInfo.Manufacturer,
                ["displayVersion"] = appInfo.Version,
                ["installCommandLine"] = installCommand,
                ["uninstallCommandLine"] = uninstallCommand,
                ["applicableArchitectures"] = "x64",
                ["fileName"] = intuneWinInfo.FileName,
                ["setupFilePath"] = "Deploy-Application.exe",
                ["installExperience"] = new Dictionary<string, object>
                {
                    ["runAsAccount"] = installContext,
                    ["deviceRestartBehavior"] = "allow"
                },
                ["detectionRules"] = formattedDetectionRules.ToArray(),
                ["returnCodes"] = new[]
                {
                    new Dictionary<string, object> { ["returnCode"] = 0, ["type"] = "success" },
                    new Dictionary<string, object> { ["returnCode"] = 3010, ["type"] = "softReboot" },
                    new Dictionary<string, object> { ["returnCode"] = 1641, ["type"] = "hardReboot" },
                    new Dictionary<string, object> { ["returnCode"] = 1618, ["type"] = "retry" }
                }
            };

            var json = JsonSerializer.Serialize(createAppPayload, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient!.PostAsync("https://graph.microsoft.com/beta/deviceAppManagement/mobileApps", content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to create Win32 app. Status: {response.StatusCode}, Response: {responseText}");
            }

            var createdApp = JsonSerializer.Deserialize<JsonElement>(responseText);
            var appId = createdApp.GetProperty("id").GetString();

            Debug.WriteLine($"✓ Created Win32 app in Intune. ID: {appId}");
            return appId ?? throw new Exception("App ID not returned from creation");
        }

        private async Task<string> CreateContentVersionAsync(string appId)
        {
            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}/microsoft.graph.win32LobApp/contentVersions";
            var content = new StringContent("{}", Encoding.UTF8, "application/json");
            var response = await _httpClient!.PostAsync(url, content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to create content version. Status: {response.StatusCode}, Response: {responseText}");
            }

            var contentVersion = JsonSerializer.Deserialize<JsonElement>(responseText);
            var contentVersionId = contentVersion.GetProperty("id").GetString();

            Debug.WriteLine($"✓ Created content version. ID: {contentVersionId}");
            return contentVersionId ?? throw new Exception("Content version ID not returned");
        }

        private async Task<string> CreateFileEntryAsync(string appId, string contentVersionId, IntuneWinInfo intuneWinInfo)
        {
            var encryptedSize = new FileInfo(intuneWinInfo.EncryptedFilePath).Length;

            var fileBody = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.mobileAppContentFile",
                ["name"] = intuneWinInfo.FileName,
                ["size"] = intuneWinInfo.UnencryptedContentSize,
                ["sizeEncrypted"] = encryptedSize,
                ["manifest"] = (object?)null,
                ["isDependency"] = false
            };

            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}/microsoft.graph.win32LobApp/contentVersions/{contentVersionId}/files";
            var json = JsonSerializer.Serialize(fileBody, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient!.PostAsync(url, content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Failed to create file entry. Status: {response.StatusCode}, Response: {responseText}");
            }

            var fileEntry = JsonSerializer.Deserialize<JsonElement>(responseText);
            var fileId = fileEntry.GetProperty("id").GetString();

            Debug.WriteLine($"✓ Created file entry. ID: {fileId}");
            return fileId ?? throw new Exception("File ID not returned");
        }

        private async Task<AzureStorageInfo> WaitForAzureStorageUriAsync(string appId, string contentVersionId, string fileId)
        {
            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}/microsoft.graph.win32LobApp/contentVersions/{contentVersionId}/files/{fileId}";
            
            for (int attempts = 0; attempts < 60; attempts++)
            {
                var response = await _httpClient!.GetAsync(url);
                var responseText = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Failed to get file info. Status: {response.StatusCode}, Response: {responseText}");
                }

                var fileInfo = JsonSerializer.Deserialize<JsonElement>(responseText);
                var uploadState = fileInfo.GetProperty("uploadState").GetString();

                if (uploadState == "AzureStorageUriRequestSuccess")
                {
                    var azureStorageUri = fileInfo.GetProperty("azureStorageUri").GetString();
                    Debug.WriteLine($"✓ Got Azure Storage URI");
                    return new AzureStorageInfo { SasUri = azureStorageUri ?? throw new Exception("Azure Storage URI is null") };
                }

                if (uploadState != "AzureStorageUriRequestPending")
                {
                    throw new Exception($"Unexpected upload state: {uploadState}");
                }

                await Task.Delay(10000); // Wait 10 seconds
            }

            throw new Exception("Timeout waiting for Azure Storage URI");
        }

        private async Task UploadFileToAzureStorageAsync(string sasUri, string filePath, IUploadProgress? progress = null)
        {
            const int chunkSize = 1024 * 1024 * 6; // 6MB chunks
            var fileInfo = new FileInfo(filePath);
            var totalSize = fileInfo.Length;
            var totalChunks = (int)Math.Ceiling((double)totalSize / chunkSize);

            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var blockIds = new List<string>();

            for (int chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++)
            {
                var blockId = Convert.ToBase64String(Encoding.ASCII.GetBytes(chunkIndex.ToString("0000")));
                blockIds.Add(blockId);

                var buffer = new byte[Math.Min(chunkSize, totalSize - (chunkIndex * chunkSize))];
                await fileStream.ReadAsync(buffer, 0, buffer.Length);

                var chunkUri = $"{sasUri}&comp=block&blockid={blockId}";
                
                using var chunkContent = new ByteArrayContent(buffer);
                chunkContent.Headers.Add("x-ms-blob-type", "BlockBlob");

                var response = await _httpClient!.PutAsync(chunkUri, chunkContent);
                if (!response.IsSuccessStatusCode)
                {
                    var errorText = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Failed to upload chunk {chunkIndex}. Status: {response.StatusCode}, Response: {errorText}");
                }

                var progressPercentage = 65 + (int)((chunkIndex + 1.0) / totalChunks * 15); // 65-80% range
                progress?.UpdateProgress(progressPercentage, $"Uploading chunk {chunkIndex + 1} of {totalChunks}...");

                Debug.WriteLine($"✓ Uploaded chunk {chunkIndex + 1}/{totalChunks}");
            }

            // Finalize the upload by committing the block list
            var blockListXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><BlockList>";
            foreach (var blockId in blockIds)
            {
                blockListXml += $"<Latest>{blockId}</Latest>";
            }
            blockListXml += "</BlockList>";

            var finalizeUri = $"{sasUri}&comp=blocklist";
            var finalizeContent = new StringContent(blockListXml, Encoding.UTF8, "application/xml");
            var finalizeResponse = await _httpClient!.PutAsync(finalizeUri, finalizeContent);

            if (!finalizeResponse.IsSuccessStatusCode)
            {
                var errorText = await finalizeResponse.Content.ReadAsStringAsync();
                throw new Exception($"Failed to finalize upload. Status: {finalizeResponse.StatusCode}, Response: {errorText}");
            }

            Debug.WriteLine($"✓ Successfully uploaded file to Azure Storage");
        }

        private async Task CommitFileAsync(string appId, string contentVersionId, string fileId, EncryptionInfo encryptionInfo)
        {
            var commitBody = new Dictionary<string, object>
            {
                ["fileEncryptionInfo"] = new Dictionary<string, object>
                {
                    ["encryptionKey"] = encryptionInfo.EncryptionKey,
                    ["macKey"] = encryptionInfo.MacKey,
                    ["initializationVector"] = encryptionInfo.InitializationVector,
                    ["mac"] = encryptionInfo.Mac,
                    ["profileIdentifier"] = encryptionInfo.ProfileIdentifier,
                    ["fileDigest"] = encryptionInfo.FileDigest,
                    ["fileDigestAlgorithm"] = encryptionInfo.FileDigestAlgorithm
                }
            };

            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}/microsoft.graph.win32LobApp/contentVersions/{contentVersionId}/files/{fileId}/commit";
            var json = JsonSerializer.Serialize(commitBody, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient!.PostAsync(url, content);

            if (!response.IsSuccessStatusCode)
            {
                var responseText = await response.Content.ReadAsStringAsync();
                throw new Exception($"Failed to commit file. Status: {response.StatusCode}, Response: {responseText}");
            }

            Debug.WriteLine($"✓ File committed successfully");
        }

        private async Task WaitForFileProcessingAsync(string appId, string contentVersionId, string fileId, string stage)
        {
            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}/microsoft.graph.win32LobApp/contentVersions/{contentVersionId}/files/{fileId}";
            var successState = $"{stage}Success";
            var pendingState = $"{stage}Pending";
            
            for (int attempts = 0; attempts < 120; attempts++)
            {
                var response = await _httpClient!.GetAsync(url);
                var responseText = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Failed to get file processing status. Status: {response.StatusCode}, Response: {responseText}");
                }

                var fileInfo = JsonSerializer.Deserialize<JsonElement>(responseText);
                var uploadState = fileInfo.GetProperty("uploadState").GetString();

                if (uploadState == successState)
                {
                    Debug.WriteLine($"✓ File processing completed for stage: {stage}");
                    return;
                }

                if (uploadState != pendingState)
                {
                    throw new Exception($"File processing failed. State: {uploadState}");
                }

                await Task.Delay(5000); // Wait 5 seconds
            }

            throw new Exception($"Timeout waiting for file processing stage: {stage}");
        }

        private async Task CommitAppAsync(string appId, string contentVersionId)
        {
            var commitBody = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobApp",
                ["committedContentVersion"] = contentVersionId
            };

            var url = $"https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/{appId}";
            var json = JsonSerializer.Serialize(commitBody, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _httpClient!.PatchAsync(url, content);

            if (!response.IsSuccessStatusCode)
            {
                var responseText = await response.Content.ReadAsStringAsync();
                throw new Exception($"Failed to commit app. Status: {response.StatusCode}, Response: {responseText}");
            }

            Debug.WriteLine($"✓ App committed successfully");
        }

        private void CleanupTempFiles(IntuneWinInfo intuneWinInfo)
        {
            try
            {
                if (Directory.Exists(intuneWinInfo.TempDirectory))
                {
                    Directory.Delete(intuneWinInfo.TempDirectory, true);
                    Debug.WriteLine($"✓ Cleaned up temp files: {intuneWinInfo.TempDirectory}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"⚠ Failed to cleanup temp files: {ex.Message}");
            }
        }

        private Dictionary<string, object>? ConvertDetectionRuleForBetaAPI(DetectionRule rule)
        {
            switch (rule.Type)
            {
                case DetectionRuleType.File:
                    return ConvertFileDetectionForBetaAPI(rule);
                    
                case DetectionRuleType.Registry:
                    return ConvertRegistryDetectionForBetaAPI(rule);
                    
                default:
                    Debug.WriteLine($"⚠ Cannot convert {rule.Type} detection rule for beta API yet");
                    return null;
            }
        }

        private Dictionary<string, object> ConvertFileDetectionForBetaAPI(DetectionRule rule)
        {
            var fileRule = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobAppFileSystemDetection",
                ["path"] = rule.Path.Trim(),
                ["fileOrFolderName"] = rule.FileOrFolderName.Trim(),
                ["check32BitOn64System"] = false
            };

            if (rule.CheckVersion && !string.IsNullOrEmpty(rule.DetectionValue))
            {
                fileRule["detectionType"] = "version";
                fileRule["operator"] = string.IsNullOrEmpty(rule.Operator) ? "greaterThanOrEqual" : rule.Operator;
                fileRule["detectionValue"] = rule.DetectionValue.Trim();
            }
            else
            {
                fileRule["detectionType"] = "exists";
            }

            return fileRule;
        }

        private Dictionary<string, object> ConvertRegistryDetectionForBetaAPI(DetectionRule rule)
        {
            var registryRule = new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.win32LobAppRegistryDetection",
                ["keyPath"] = rule.RegistryKey.Trim(),
                ["check32BitOn64System"] = false
            };

            if (!string.IsNullOrEmpty(rule.RegistryValueName))
            {
                registryRule["valueName"] = rule.RegistryValueName.Trim();
                
                if (!string.IsNullOrEmpty(rule.ExpectedValue))
                {
                    registryRule["detectionType"] = "string";
                    registryRule["operator"] = string.IsNullOrEmpty(rule.Operator) ? "equal" : rule.Operator;
                    registryRule["detectionValue"] = rule.ExpectedValue.Trim();
                }
                else
                {
                    registryRule["detectionType"] = "exists";
                }
            }
            else
            {
                registryRule["detectionType"] = "exists";
            }

            return registryRule;
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    
}