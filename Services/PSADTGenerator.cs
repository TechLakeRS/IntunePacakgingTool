using System.IO;
using System.Threading.Tasks;
using System;
using System.Linq;

namespace IntunePackagingTool
{
    public class PSADTGenerator
    {
        private readonly string _baseOutputPath = @"\\nbb.local\sys\SCCMData\IntuneApplications";
        private readonly string _templatePath = @"\\nbb.local\sys\SCCMData\TOOLS\IntunePackagingTool\20250811\Application";

        public async Task<string> CreatePackageAsync(ApplicationInfo appInfo)
        {
            try
            {
                // Create folder structure: Manufacturer_AppName\Version\ (replace spaces with underscores)
                var cleanManufacturer = appInfo.Manufacturer.Replace(" ", "_");
                var cleanAppName = appInfo.Name.Replace(" ", "_");
                var appFolderName = $"{cleanManufacturer}_{cleanAppName}";
                var appBasePath = Path.Combine(_baseOutputPath, appFolderName);
                var packagePath = Path.Combine(appBasePath, appInfo.Version);

                // Check if Manufacturer_Appname already exists
                if (!Directory.Exists(appBasePath))
                {
                    Directory.CreateDirectory(appBasePath);
                    Console.WriteLine($"Created new application folder: {appBasePath}");
                }
                else
                {
                    Console.WriteLine($"Application folder already exists: {appBasePath}");
                }

                // Check if version folder already exists
                if (Directory.Exists(packagePath))
                {
                    throw new InvalidOperationException($"Version {appInfo.Version} already exists for {cleanManufacturer}_{cleanAppName}");
                }

                // Create version folder
                Directory.CreateDirectory(packagePath);
                Console.WriteLine($"Created version folder: {packagePath}");

                // Create required subfolders
                var folders = new[] { "Application", "Documentation", "Icon", "Intune", "NBB_Info", "Project Files", "Sources" };
                foreach (var folder in folders)
                {
                    var fullPath = Path.Combine(packagePath, folder);
                    Directory.CreateDirectory(fullPath);
                }

                // Copy template files from archive to Application folder
                await CopyTemplateFilesAsync(packagePath);

                // Copy source files to Application\Files if provided
                if (!string.IsNullOrWhiteSpace(appInfo.SourcesPath))
                {
                    // Check if it's multiple files (semicolon separated), single file, or directory
                    bool hasValidSource = false;
                    
                    if (appInfo.SourcesPath.Contains(";"))
                    {
                        // Multiple files - check if at least one exists
                        var filePaths = appInfo.SourcesPath.Split(';', StringSplitOptions.RemoveEmptyEntries);
                        hasValidSource = filePaths.Any(path => File.Exists(path.Trim()));
                    }
                    else
                    {
                        // Single file or directory
                        hasValidSource = File.Exists(appInfo.SourcesPath) || Directory.Exists(appInfo.SourcesPath);
                    }
                    
                    if (hasValidSource)
                    {
                        await CopySourceFilesToApplicationAsync(appInfo.SourcesPath, packagePath);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Source path not found or invalid: {appInfo.SourcesPath}");
                    }
                }

                // Modify the Deploy-Application.ps1 in the Application folder with metadata
                await ModifyPSADTScriptAsync(packagePath, appInfo);

                return packagePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error creating package: {ex.Message}", ex);
            }
        }

        private async Task CopyTemplateFilesAsync(string packagePath)
        {
            try
            {
                var destinationAppFolder = Path.Combine(packagePath, "Application");
                
                if (!Directory.Exists(_templatePath))
                {
                    throw new DirectoryNotFoundException($"Template path not found: {_templatePath}");
                }

                await Task.Run(() =>
                {
                    // Copy all files and folders from template to Application folder
                    CopyDirectory(_templatePath, destinationAppFolder, true);
                });

                Console.WriteLine($"Copied template files from {_templatePath} to {destinationAppFolder}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error copying template files: {ex.Message}", ex);
            }
        }

        private async Task CopySourceFilesToApplicationAsync(string sourcePath, string packagePath)
        {
            try
            {
                var applicationFilesPath = Path.Combine(packagePath, "Application", "Files");
                
                // Create Application\Files folder if it doesn't exist
                Directory.CreateDirectory(applicationFilesPath);

                await Task.Run(() =>
                {
                    // Check if sourcePath contains multiple files (semicolon separated) or single file/directory
                    if (sourcePath.Contains(";"))
                    {
                        // Handle multiple files
                        var filePaths = sourcePath.Split(';', StringSplitOptions.RemoveEmptyEntries);
                        
                        foreach (var filePath in filePaths)
                        {
                            var trimmedPath = filePath.Trim();
                            if (File.Exists(trimmedPath))
                            {
                                var fileName = Path.GetFileName(trimmedPath);
                                var destinationFile = Path.Combine(applicationFilesPath, fileName);
                                File.Copy(trimmedPath, destinationFile, true);
                                Console.WriteLine($"Copied file: {fileName}");
                            }
                        }
                    }
                    else if (File.Exists(sourcePath))
                    {
                        // Handle single file
                        var fileName = Path.GetFileName(sourcePath);
                        var destinationFile = Path.Combine(applicationFilesPath, fileName);
                        File.Copy(sourcePath, destinationFile, true);
                        Console.WriteLine($"Copied file: {fileName}");
                    }
                    else if (Directory.Exists(sourcePath))
                    {
                        // Handle directory (existing functionality)
                        foreach (var file in Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories))
                        {
                            var relativePath = Path.GetRelativePath(sourcePath, file);
                            var destinationFile = Path.Combine(applicationFilesPath, relativePath);
                            
                            var destinationDir = Path.GetDirectoryName(destinationFile);
                            if (!string.IsNullOrEmpty(destinationDir))
                            {
                                Directory.CreateDirectory(destinationDir);
                            }
                            File.Copy(file, destinationFile, true);
                        }
                        Console.WriteLine($"Copied all files from directory: {sourcePath}");
                    }
                    else
                    {
                        throw new FileNotFoundException($"Source path not found: {sourcePath}");
                    }
                });

                Console.WriteLine($"Copied source files to {applicationFilesPath}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error copying source files: {ex.Message}", ex);
            }
        }

        private void CopyDirectory(string sourceDir, string destinationDir, bool recursive)
        {
            // Get the subdirectory info
            var dir = new DirectoryInfo(sourceDir);

            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");

            // Cache directories before we start copying
            DirectoryInfo[] dirs = dir.GetDirectories();

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the directory and copy them to the new location
            foreach (FileInfo file in dir.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath, true);
            }

            // If we are copying subdirectories, copy them and their contents to the new location
            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
            }
        }

        private async Task ModifyPSADTScriptAsync(string packagePath, ApplicationInfo appInfo)
        {
            try
            {
                var scriptPath = Path.Combine(packagePath, "Application", "Deploy-Application.ps1");
                
                if (!File.Exists(scriptPath))
                {
                    throw new FileNotFoundException($"Deploy-Application.ps1 not found in copied template: {scriptPath}");
                }

                // Read the existing script
                var scriptContent = await File.ReadAllTextAsync(scriptPath);

                // Update the metadata variables in the script
                scriptContent = UpdateScriptMetadata(scriptContent, appInfo);

                // Write the modified script back
                await File.WriteAllTextAsync(scriptPath, scriptContent);
                
                Console.WriteLine($"Modified Deploy-Application.ps1 with metadata: {scriptPath}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error modifying PSADT script: {ex.Message}", ex);
            }
        }

        private string UpdateScriptMetadata(string scriptContent, ApplicationInfo appInfo)
        {
            // Update the variable declarations section with user's metadata
            var lines = scriptContent.Split('\n').ToList();
            
            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                
                // Update specific variables (handle both [string] and [String], with or without values)
                if (line.Contains("$appVendor") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appVendor = '{appInfo.Manufacturer}'";
                }
                else if (line.Contains("$appName") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appName = '{appInfo.Name}'";
                }
                else if (line.Contains("$appVersion") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appVersion = '{appInfo.Version}'";
                }
                else if (line.Contains("$appArch") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appArch = 'x64'";
                }
                else if (line.Contains("$appScriptDate") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appScriptDate = '{DateTime.Now:yyyy/MM/dd}'";
                }
                else if (line.Contains("$appScriptAuthor") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$appScriptAuthor = 'NBB Application Packaging Tools'";
                }
                // Add ServiceNow SRI if it doesn't exist, or update if it does
                else if (line.Contains("$ServiceNowSRI") && (line.StartsWith("[string]") || line.StartsWith("[String]")))
                {
                    lines[i] = $"\t[string]$ServiceNowSRI = '{appInfo.ServiceNowSRI}'";
                }
                // If we're after the appScriptAuthor line and ServiceNowSRI doesn't exist, add it
                else if (line.Contains("$appScriptAuthor") && !scriptContent.Contains("$ServiceNowSRI"))
                {
                    // Insert ServiceNow SRI after appScriptAuthor
                    lines.Insert(i + 1, $"\t[string]$ServiceNowSRI = '{appInfo.ServiceNowSRI}'");
                    break; // Exit loop to avoid infinite insertion
                }
            }
            
            return string.Join('\n', lines);
        }

        public string GetScriptPath(string manufacturer, string appName, string version)
        {
            // Script is now in the Application folder, not the root
            var cleanManufacturer = manufacturer.Replace(" ", "_");
            var cleanAppName = appName.Replace(" ", "_");
            return Path.Combine(_baseOutputPath, $"{cleanManufacturer}_{cleanAppName}", version, "Application", "Deploy-Application.ps1");
        }
    }
}