using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace IntunePackagingTool
{
    public partial class AddFileDetectionDialog : Window
    {
        public DetectionRule? DetectionRule { get; private set; }

        public AddFileDetectionDialog()
        {
            InitializeComponent();
            UpdatePreview();
        }

        private void BrowsePathButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select any file in the target folder",
                Filter = "All files (*.*)|*.*",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Select folder"
            };

            if (dialog.ShowDialog() == true)
            {
                var selectedPath = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(selectedPath))
                {
                    FilePathTextBox.Text = selectedPath;
                    UpdatePreview();
                }
            }
        }

        private void DetectionTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ValueLabel == null || ValuePanel == null) return;

            var selectedItem = (ComboBoxItem)DetectionTypeCombo.SelectedItem;
            var content = selectedItem?.Content?.ToString() ?? "";

            if (content == "File or folder exists")
            {
                ValueLabel.Visibility = Visibility.Collapsed;
                ValuePanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ValueLabel.Visibility = Visibility.Visible;
                ValuePanel.Visibility = Visibility.Visible;

                ValueLabel.Text = content switch
                {
                    "File version" => "Version:",
                    "File size (MB)" => "Size (MB):",
                    "Date modified" => "Date:",
                    _ => "Value:"
                };
            }

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTextBlock == null) return;

            var path = FilePathTextBox?.Text ?? "";
            var fileName = FileNameTextBox?.Text ?? "";
            var fullPath = Path.Combine(path, fileName);

            var selectedType = (ComboBoxItem)DetectionTypeCombo?.SelectedItem;
            var typeContent = selectedType?.Content?.ToString() ?? "File or folder exists";

            if (typeContent == "File or folder exists")
            {
                PreviewTextBlock.Text = $"Check if file exists: {fullPath}";
            }
            else
            {
                var operatorItem = (ComboBoxItem)OperatorCombo?.SelectedItem;
                var operatorText = operatorItem?.Content?.ToString() ?? "Equal to";
                var value = ValueTextBox?.Text ?? "";
                
                PreviewTextBlock.Text = $"Check if file exists: {fullPath}\nAND {typeContent.ToLower()} is {operatorText.ToLower()} {value}";
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(FilePathTextBox.Text))
            {
                MessageBox.Show("Please enter a file path.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(FileNameTextBox.Text))
            {
                MessageBox.Show("Please enter a file name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DetectionRule = new DetectionRule
            {
                Type = DetectionRuleType.File,
                Path = FilePathTextBox.Text.Trim(),
                FileOrFolderName = FileNameTextBox.Text.Trim()
            };

            var selectedType = (ComboBoxItem)DetectionTypeCombo.SelectedItem;
            var typeContent = selectedType?.Content?.ToString() ?? "";

            if (typeContent != "File or folder exists")
            {
                DetectionRule.CheckVersion = typeContent == "File version";
                DetectionRule.DetectionValue = ValueTextBox.Text.Trim();
                
                var operatorItem = (ComboBoxItem)OperatorCombo.SelectedItem;
                DetectionRule.Operator = operatorItem?.Tag?.ToString() ?? "equal";
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}