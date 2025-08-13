using System;
using System.Windows;
using System.Windows.Controls;

namespace IntunePackagingTool
{
    public partial class AddRegistryDetectionDialog : Window
    {
        public DetectionRule? DetectionRule { get; private set; }

        public AddRegistryDetectionDialog()
        {
            InitializeComponent();
            UpdatePreview();
        }

        private void DetectionTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ExpectedValueLabel == null || ExpectedValuePanel == null) return;

            var selectedItem = (ComboBoxItem)DetectionTypeCombo.SelectedItem;
            var content = selectedItem?.Content?.ToString() ?? "";

            if (content == "Key or value exists")
            {
                ExpectedValueLabel.Visibility = Visibility.Collapsed;
                ExpectedValuePanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ExpectedValueLabel.Visibility = Visibility.Visible;
                ExpectedValuePanel.Visibility = Visibility.Visible;
            }

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTextBlock == null) return;

            var rootKey = ((ComboBoxItem)RootKeyCombo?.SelectedItem)?.Content?.ToString() ?? "HKEY_LOCAL_MACHINE";
            var keyPath = KeyPathTextBox?.Text ?? "";
            var valueName = ValueNameTextBox?.Text ?? "";
            var fullPath = $"{rootKey}\\{keyPath}\\{valueName}";

            var selectedType = (ComboBoxItem)DetectionTypeCombo?.SelectedItem;
            var typeContent = selectedType?.Content?.ToString() ?? "Key or value exists";

            if (typeContent == "Key or value exists")
            {
                PreviewTextBlock.Text = $"Registry key exists: {fullPath}";
            }
            else
            {
                var operatorItem = (ComboBoxItem)OperatorCombo?.SelectedItem;
                var operatorText = operatorItem?.Content?.ToString() ?? "Equal to";
                var expectedValue = ExpectedValueTextBox?.Text ?? "";
                
                PreviewTextBlock.Text = $"Registry value: {fullPath}\n{typeContent} {operatorText.ToLower()} {expectedValue}";
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(KeyPathTextBox.Text))
            {
                MessageBox.Show("Please enter a registry key path.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DetectionRule = new DetectionRule
            {
                Type = DetectionRuleType.Registry,
                RegistryHive = ((ComboBoxItem)RootKeyCombo.SelectedItem)?.Content?.ToString() ?? "HKEY_LOCAL_MACHINE",
                RegistryKey = KeyPathTextBox.Text.Trim(),
                RegistryValueName = ValueNameTextBox.Text.Trim()
            };

            var selectedType = (ComboBoxItem)DetectionTypeCombo.SelectedItem;
            var typeContent = selectedType?.Content?.ToString() ?? "";

            if (typeContent != "Key or value exists")
            {
                DetectionRule.ExpectedValue = ExpectedValueTextBox.Text.Trim();
                
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