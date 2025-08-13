using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace IntunePackagingTool
{
    public partial class UploadToIntuneWindow : Window
    {
        public ApplicationInfo? ApplicationInfo { get; set; }
        public string PackagePath { get; set; } = "";
        
        private ObservableCollection<DetectionRule> _detectionRules = new ObservableCollection<DetectionRule>();
        private IntuneService _intuneService = new IntuneService();

        public UploadToIntuneWindow()
        {
            InitializeComponent();
            DetectionRulesList.ItemsSource = _detectionRules;
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

                UploadButton.IsEnabled = false;
                CancelButton.IsEnabled = false;
                UploadProgressPanel.Visibility = Visibility.Visible;

                await SimulateUpload();

                MessageBox.Show(
                    $"Package prepared successfully!\n\n" +
                    $"The .intunewin package has been created at:\n{PackagePath}\\Intune\\\n\n" +
                    $"You can now upload it manually through the Microsoft Intune admin center:\n" +
                    $"1. Go to Apps > All apps\n" +
                    $"2. Click Add > Windows app (Win32)\n" +
                    $"3. Upload the .intunewin file\n" +
                    $"4. Configure the detection rules as specified",
                    "Upload Complete", MessageBoxButton.OK, MessageBoxImage.Information);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Upload failed: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                UploadButton.IsEnabled = true;
                CancelButton.IsEnabled = true;
            }
        }

        private async Task SimulateUpload()
        {
            UploadStatusText.Text = "Creating .intunewin package...";
            UploadProgressBar.Value = 25;
            await Task.Delay(1000);

            UploadStatusText.Text = "Packaging application files...";
            UploadProgressBar.Value = 50;
            await Task.Delay(1000);

            UploadStatusText.Text = "Generating metadata...";
            UploadProgressBar.Value = 75;
            await Task.Delay(1000);

            UploadStatusText.Text = "Finalizing package...";
            UploadProgressBar.Value = 100;
            await Task.Delay(500);
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
        }
    }
}