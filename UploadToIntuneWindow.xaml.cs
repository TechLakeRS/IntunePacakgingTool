using System;
using System.Collections.ObjectModel;
using System.Diagnostics;  
using System.IO;           
using System.Linq;         
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace IntunePackagingTool
{
    public partial class UploadToIntuneWindow : Window, IUploadProgress
    {
        public ApplicationInfo? ApplicationInfo { get; set; }
        public string PackagePath { get; set; } = "";
        
        private ObservableCollection<DetectionRule> _detectionRules = new ObservableCollection<DetectionRule>();
        private IntuneService _intuneService = new IntuneService();
        private IntuneUploadService _uploadService;

        public UploadToIntuneWindow()
        {
            InitializeComponent();
            DetectionRulesList.ItemsSource = _detectionRules;
            _uploadService = new IntuneUploadService(_intuneService);
        }

        // Implement IUploadProgress interface
        public void UpdateProgress(int percentage, string message)
        {
            Dispatcher.Invoke(() =>
            {
                UploadProgressBar.Value = percentage;
                UploadStatusText.Text = message;
            });
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            
            if (ApplicationInfo != null)
            {
                AppSummaryText.Text = $"{ApplicationInfo.Manufacturer} {ApplicationInfo.Name} v{ApplicationInfo.Version}";
                DescriptionTextBox.Text = $"{ApplicationInfo.Name} packaged with NBB PSADT Tools";
            }

            // Add a default file detection rule
            _detectionRules.Add(new DetectionRule
            {
                Type = DetectionRuleType.File,
                Path = "%ProgramFiles%",
                FileOrFolderName = $"{ApplicationInfo?.Name ?? "MyApp"}.exe",
                CheckVersion = false
            });
        }

        private void AddFileDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddFileDetectionDialog
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true && dialog.DetectionRule != null)
            {
                _detectionRules.Add(dialog.DetectionRule);
            }
        }

        private void AddRegistryDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddRegistryDetectionDialog
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true && dialog.DetectionRule != null)
            {
                _detectionRules.Add(dialog.DetectionRule);
            }
        }

        private void AddScriptDetectionButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Script detection rule editor will be implemented in a future version.", 
                "Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BrowseIconButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Application Icon",
                Filter = "Image files (*.png;*.ico;*.jpg)|*.png;*.ico;*.jpg|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                DefaultIconText.Text = "✅";
                MessageBox.Show($"Icon selected: {System.IO.Path.GetFileName(dialog.FileName)}", 
                    "Icon Selected", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void RemoveDetectionRule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is DetectionRule ruleToRemove)
            {
                var result = MessageBox.Show(
                    $"Are you sure you want to remove this detection rule?\n\n{ruleToRemove.Title}", 
                    "Remove Detection Rule", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);
                    
                if (result == MessageBoxResult.Yes)
                {
                    _detectionRules.Remove(ruleToRemove);
                }
            }
        }

        private async void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_detectionRules.Count == 0)
                {
                    MessageBox.Show("Please add at least one detection rule.", "Validation Error", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("=== UI DETECTION RULES ===");
                System.Diagnostics.Debug.WriteLine($"Total rules from UI: {_detectionRules.Count}");
                foreach (var rule in _detectionRules)
                {
                    System.Diagnostics.Debug.WriteLine($"Rule: Type={rule.Type}, Path='{rule.Path}', File='{rule.FileOrFolderName}', CheckVersion={rule.CheckVersion}");
                }
                System.Diagnostics.Debug.WriteLine("=== END UI DETECTION RULES ===");

                if (string.IsNullOrEmpty(PackagePath) || ApplicationInfo == null)
                {
                    MessageBox.Show("Package information is not available.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                UploadButton.IsEnabled = false;
                CancelButton.IsEnabled = false;
                UploadProgressPanel.Visibility = Visibility.Visible;

                // Get form values INCLUDING install context
                var installCommand = InstallCommandTextBox.Text.Trim();
                var uninstallCommand = UninstallCommandTextBox.Text.Trim();
                var description = DescriptionTextBox.Text.Trim();
                
                // Read the install context from the UI
                var selectedInstallContext = ((ComboBoxItem)InstallContextCombo.SelectedItem)?.Content?.ToString() ?? "System";
                var installContext = selectedInstallContext.ToLower(); // "system" or "user"

                // Use the upload service with install context
                var appId = await _uploadService.UploadWin32ApplicationAsync(
                    ApplicationInfo, 
                    PackagePath, 
                    _detectionRules.ToList(), 
                    installCommand, 
                    uninstallCommand, 
                    description,
                    installContext,  // Pass the actual UI selection
                    this);

                var intuneFolder = System.IO.Path.Combine(PackagePath, "Intune");
                var intuneWinFiles = System.IO.Directory.GetFiles(intuneFolder, "*.intunewin");
                var intuneWinFile = intuneWinFiles.Length > 0 ? intuneWinFiles[0] : "";

                MessageBox.Show(
                    $"🎉 Application successfully uploaded to Microsoft Intune!\n\n" +
                    $"Application: {ApplicationInfo.Manufacturer} {ApplicationInfo.Name} v{ApplicationInfo.Version}\n" +
                    $"Intune App ID: {appId}\n\n" +
                    $"Package Details:\n" +
                    $"• Install Context: {selectedInstallContext}\n" +
                    $"• Install Command: {installCommand}\n" +
                    $"• Uninstall Command: {uninstallCommand}\n" +
                    $"• Detection Rules: {_detectionRules.Count} configured\n\n" +
                    $"The application is now available in the Microsoft Intune admin center.\n" +
                    $"Local .intunewin file: {System.IO.Path.GetFileName(intuneWinFile)}",
                    "Upload Successful!", MessageBoxButton.OK, MessageBoxImage.Information);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                UploadProgressPanel.Visibility = Visibility.Collapsed;
                MessageBox.Show(
                    $"Upload failed: {ex.Message}\n\n" +
                    $"The .intunewin file may have been created locally, but the upload to Intune failed.\n" +
                    $"You can upload it manually through the Intune admin center.", 
                    "Upload Error", MessageBoxButton.OK, MessageBoxImage.Error);
                
                Debug.WriteLine($"Upload error: {ex}");
            }
            finally
            {
                UploadButton.IsEnabled = true;
                CancelButton.IsEnabled = true;
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _intuneService?.Dispose();
            _uploadService?.Dispose();
        }
    }
}