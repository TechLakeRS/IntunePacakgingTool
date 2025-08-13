using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using System.Text.Json;
using System.Net.Http;
using System.Text;
using System.Diagnostics;

namespace IntunePackagingTool
{
    public partial class UploadToIntuneWindow : Window
    {
        private ApplicationInfo _appInfo;
        private string _packagePath;
        private IntuneService _intuneService;
        private ObservableCollection<DetectionRule> _detectionRules;
        private string? _selectedIconPath;

        public UploadToIntuneWindow(ApplicationInfo appInfo, string packagePath, IntuneService intuneService)
        {
            InitializeComponent();
            
            _appInfo = appInfo;
            _packagePath = packagePath;
            _intuneService = intuneService;
            _detectionRules = new ObservableCollection<DetectionRule>();
            
            InitializeWindow();
        }

        private void InitializeWindow()
        {
            // Update window title and summary
            AppSummaryText.Text = $"{_appInfo.Manufacturer} {_appInfo.Name} v{_appInfo.Version}";
            
            // Bind detection rules to the list
            DetectionRulesList.ItemsSource = _detectionRules;
            
            // Add a default file detection rule
            AddDefaultDetectionRule();
        }

        private void AddDefaultDetectionRule()
        {
            // Add a basic file detection for Deploy-Application.exe
            var defaultRule = new FileDetectionRule
            {
                Icon = "📁",
                Title = "PSADT Detection",
                Description = $"File exists: Deploy-Application.exe",
                Path = "%ProgramFiles%",
                FileName = "Deploy-Application.exe",
                DetectionType = FileDetectionType.Exists
            };
            
            _detectionRules.Add(defaultRule);
        }

        #region Detection Rule Management

        private void AddFileDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddFileDetectionDialog();
            if (dialog.ShowDialog() == true)
            {
                _detectionRules.Add(dialog.DetectionRule);
            }
        }

        private void AddRegistryDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddRegistryDetectionDialog();
            if (dialog.ShowDialog() == true)
            {
                _detectionRules.Add(dialog.DetectionRule);
            }
        }

        private void AddScriptDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddScriptDetectionDialog();
            if (dialog.ShowDialog() == true)
            {
                _detectionRules.Add(dialog.DetectionRule);
            }
        }

        #endregion

        #region Icon Management

        private void BrowseIconButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Title = "Select Application Icon",
                Filter = "Image Files (*.png;*.ico;*.jpg;*.jpeg)|*.png;*.ico;*.jpg;*.jpeg|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (fileDialog.ShowDialog() == true)
            {
                try
                {
                    _selectedIconPath = fileDialog.FileName;
                    
                    // Load and display the icon
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(_selectedIconPath);
                    bitmap.DecodePixelWidth = 64;
                    bitmap.EndInit();
                    
                    AppIconImage.Source = bitmap;
                    DefaultIconText.Visibility = Visibility.Collapsed;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading icon: {ex.Message}", "Icon Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        #endregion

        #region Upload Process

        private async void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Disable the upload button
                UploadButton.IsEnabled = false;
                UploadProgressPanel.Visibility = Visibility.Visible;

                // Step 1: Create .intunewin file
                UpdateProgress(10, "Creating .intunewin package...");
                var intuneWinPath = await CreateIntuneWinFileAsync();

                // Step 2: Upload to Intune
                UpdateProgress(30, "Uploading to Microsoft Intune...");
                await UploadToIntuneAsync(intuneWinPath);

                UpdateProgress(100, "Upload completed successfully!");
                
                MessageBox.Show("Application uploaded to Intune successfully!", "Upload Complete", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Upload failed: {ex.Message}", "Upload Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                
                UploadButton.IsEnabled = true;
                UploadProgressPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateProgress(int percentage, string status)
        {
            UploadProgressBar.Value = percentage;
            UploadStatusText.Text = status;
            
            // Force UI update
            Application.Current.Dispatcher.Invoke(() => { });
        }

        private async Task<string> CreateIntuneWinFileAsync()
        {
            // This would use the IntuneWinAppUtil.exe tool to create the .intunewin file
            // For now, we'll simulate this process
            
            var applicationFolderPath = Path.Combine(_packagePath, "Application");
            var outputPath = Path.Combine(_packagePath, "Intune");
            var intuneWinPath = Path.Combine(outputPath, $"{_appInfo.Manufacturer}_{_appInfo.Name}_{_appInfo.Version}.intunewin");

            // Ensure output directory exists
            Directory.CreateDirectory(outputPath);

            // In a real implementation, you would:
            // 1. Download IntuneWinAppUtil.exe if not present
            // 2. Run: IntuneWinAppUtil.exe -c [source] -s Deploy-Application.exe -o [output]
            
            // For demo purposes, create a placeholder file
            await File.WriteAllTextAsync(intuneWinPath, "Placeholder .intunewin file");
            
            await Task.Delay(2000); // Simulate processing time
            
            return intuneWinPath;
        }

        private async Task UploadToIntuneAsync(string intuneWinPath)
        {
            // Create the Win32 LOB app object
            var win32App = new
            {
                displayName = $"{_appInfo.Manufacturer} {_appInfo.Name}",
                description = DescriptionTextBox.Text,
                publisher = _appInfo.Manufacturer,
                displayVersion = _appInfo.Version,
                installCommandLine = InstallCommandTextBox.Text,
                uninstallCommandLine = UninstallCommandTextBox.Text,
                installExperience = new
                {
                    runAsAccount = InstallContextCombo.SelectedIndex == 0 ? "system" : "user"
                },
                detectionRules = _detectionRules.Select(rule => rule.ToJson()).ToArray(),
                notes = $"Created by NBB Application Packaging Tools on {DateTime.Now:yyyy-MM-dd HH:mm}"
            };

            // This would involve multiple Graph API calls:
            // 1. Create the app
            // 2. Upload the .intunewin file content
            // 3. Commit the upload
            
            // For now, simulate the upload process
            UpdateProgress(50, "Creating application in Intune...");
            await Task.Delay(1500);
            
            UpdateProgress(70, "Uploading package content...");
            await Task.Delay(2000);
            
            UpdateProgress(90, "Finalizing upload...");
            await Task.Delay(1000);
        }

        #endregion

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    #region Detection Rule Classes

    public abstract class DetectionRule
    {
        public string Icon { get; set; } = "";
        public string Title { get; set; } = "";
        public string Description { get; set; } = "";
        
        public abstract object ToJson();
    }

    public class FileDetectionRule : DetectionRule
    {
        public string Path { get; set; } = "";
        public string FileName { get; set; } = "";
        public FileDetectionType DetectionType { get; set; }
        public string Version { get; set; } = "";

        public override object ToJson()
        {
            return new
            {
                odataType = "#microsoft.graph.win32LobAppFileSystemDetection",
                path = Path,
                fileOrFolderName = FileName,
                check32BitOn64System = false,
                detectionType = DetectionType.ToString().ToLower(),
                @operator = "greaterThanOrEqual",
                detectionValue = Version
            };
        }
    }

    public class RegistryDetectionRule : DetectionRule
    {
        public string KeyPath { get; set; } = "";
        public string ValueName { get; set; } = "";
        public RegistryDetectionType DetectionType { get; set; }
        public string ExpectedValue { get; set; } = "";

        public override object ToJson()
        {
            return new
            {
                odataType = "#microsoft.graph.win32LobAppRegistryDetection",
                keyPath = KeyPath,
                valueName = ValueName,
                detectionType = DetectionType.ToString().ToLower(),
                detectionValue = ExpectedValue
            };
        }
    }

    public class ScriptDetectionRule : DetectionRule
    {
        public string ScriptContent { get; set; } = "";
        public ScriptType ScriptType { get; set; }

        public override object ToJson()
        {
            return new
            {
                odataType = "#microsoft.graph.win32LobAppPowerShellScriptDetection",
                scriptContent = Convert.ToBase64String(Encoding.UTF8.GetBytes(ScriptContent)),
                enforceSignatureCheck = false,
                runAs32Bit = false
            };
        }
    }

    public enum FileDetectionType
    {
        Exists,
        Version,
        Size,
        DateModified
    }

    public enum RegistryDetectionType
    {
        Exists,
        String,
        Integer,
        Version
    }

    public enum ScriptType
    {
        PowerShell,
        Cmd
    }

    #endregion
}