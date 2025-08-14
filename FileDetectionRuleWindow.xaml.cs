using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace IntunePackagingTool
{
    public partial class FileDetectionRuleWindow : Window
    {
        public FileDetectionRuleWindow()
        {
            InitializeComponent();
            UpdatePreview();
            
            // Wire up events for real-time preview
            PathTextBox.TextChanged += (s, e) => UpdatePreview();
            FileNameTextBox.TextChanged += (s, e) => UpdatePreview();
            VersionTextBox.TextChanged += (s, e) => UpdatePreview();
            OperatorComboBox.SelectionChanged += (s, e) => UpdatePreview();
        }

        private void FileVersionRadio_Checked(object sender, RoutedEventArgs e)
        {
            if (VersionPanel != null)
            {
                VersionPanel.Visibility = Visibility.Visible;
                UpdatePreview();
            }
        }

        private void FileVersionRadio_Unchecked(object sender, RoutedEventArgs e)
        {
            if (VersionPanel != null)
            {
                VersionPanel.Visibility = Visibility.Collapsed;
                UpdatePreview();
            }
        }

        private void UpdatePreview()
        {
            if (PreviewText == null) return;

            var path = PathTextBox?.Text ?? "";
            var fileName = FileNameTextBox?.Text ?? "";
            var fullPath = Path.Combine(path, fileName);

            if (FileExistsRadio?.IsChecked == true)
            {
                PreviewText.Text = $"Check if file exists: {fullPath}";
            }
            else if (FileVersionRadio?.IsChecked == true)
            {
                var operatorText = ((ComboBoxItem)OperatorComboBox?.SelectedItem)?.Content?.ToString() ?? "greater than or equal";
                var version = VersionTextBox?.Text ?? "1.0.0";
                PreviewText.Text = $"Check if file exists: {fullPath}\nAND file version is {operatorText.ToLower()} {version}";
            }
        }

        private void BrowsePathButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Use file dialog to browse for a folder (workaround for older .NET versions)
                var dialog = new OpenFileDialog
                {
                    Title = "Select any file in the target folder",
                    Filter = "All files (*.*)|*.*",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    FileName = "Select folder"
                };

                // Set initial directory if current path is valid
                var currentPath = PathTextBox.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(currentPath) && Directory.Exists(currentPath))
                {
                    dialog.InitialDirectory = currentPath;
                }

                if (dialog.ShowDialog() == true)
                {
                    // Get the directory from the selected file
                    var selectedPath = Path.GetDirectoryName(dialog.FileName);
                    if (!string.IsNullOrEmpty(selectedPath))
                    {
                        PathTextBox.Text = selectedPath;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting folder: {ex.Message}", "Browse Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateInput())
            {
                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool ValidateInput()
        {
            // Validate path
            if (string.IsNullOrWhiteSpace(PathTextBox.Text))
            {
                MessageBox.Show("Please enter a valid path.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                PathTextBox.Focus();
                return false;
            }

            // Validate file name
            if (string.IsNullOrWhiteSpace(FileNameTextBox.Text))
            {
                MessageBox.Show("Please enter a file or folder name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                FileNameTextBox.Focus();
                return false;
            }

            
            if (FileVersionRadio.IsChecked == true)
            {
                if (string.IsNullOrWhiteSpace(VersionTextBox.Text))
                {
                    MessageBox.Show("Please enter a version number when using version detection.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    VersionTextBox.Focus();
                    return false;
                }

                
                try
                {
                    var version = new Version(VersionTextBox.Text);
                }
                catch
                {
                    MessageBox.Show("Please enter a valid version number (e.g., 1.0.0 or 1.0.0.0).", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    VersionTextBox.Focus();
                    return false;
                }
            }

            return true;
        }

        public DetectionRule GetDetectionRule()
        {
            var rule = new DetectionRule
            {
                Type = DetectionRuleType.File,
                Path = PathTextBox.Text.Trim(),
                FileOrFolderName = FileNameTextBox.Text.Trim(),
                CheckVersion = FileVersionRadio.IsChecked == true
            };

            if (rule.CheckVersion)
            {
                rule.DetectionValue = VersionTextBox.Text.Trim();
                rule.Operator = ((ComboBoxItem)OperatorComboBox.SelectedItem)?.Tag?.ToString() ?? "greaterThanOrEqual";
            }

            return rule;
        }
    }
}
