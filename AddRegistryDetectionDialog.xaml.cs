using System;
using System.Windows;
using System.Windows.Controls;

namespace IntunePackagingTool
{
    public partial class AddRegistryDetectionDialog : Window
    {
        public RegistryDetectionRule DetectionRule { get; private set; }

        public AddRegistryDetectionDialog()
        {
            InitializeComponent();
            UpdatePreview();
            
            // Wire up events for real-time preview updates
            RootKeyCombo.SelectionChanged += (s, e) => UpdatePreview();
            KeyPathTextBox.TextChanged += (s, e) => UpdatePreview();
            ValueNameTextBox.TextChanged += (s, e) => UpdatePreview();
            ExpectedValueTextBox.TextChanged += (s, e) => UpdatePreview();
            OperatorCombo.SelectionChanged += (s, e) => UpdatePreview();
        }

        private void DetectionTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DetectionTypeCombo.SelectedIndex == 0) // Key/Value exists
            {
                ExpectedValueLabel.Visibility = Visibility.Collapsed;
                ExpectedValuePanel.Visibility = Visibility.Collapsed;
            }
            else // String, Integer, or Version value
            {
                ExpectedValueLabel.Visibility = Visibility.Visible;
                ExpectedValuePanel.Visibility = Visibility.Visible;
                
                switch (DetectionTypeCombo.SelectedIndex)
                {
                    case 1: // String
                        ExpectedValueLabel.Text = "Expected:";
                        ExpectedValueTextBox.Text = "MyApplication";
                        break;
                    case 2: // Integer
                        ExpectedValueLabel.Text = "Expected:";
                        ExpectedValueTextBox.Text = "1";
                        break;
                    case 3: // Version
                        ExpectedValueLabel.Text = "Expected:";
                        ExpectedValueTextBox.Text = "1.0.0";
                        break;
                }
            }
            
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTextBlock == null) return;

            var rootKey = ((ComboBoxItem)RootKeyCombo.SelectedItem)?.Content?.ToString() ?? "HKEY_LOCAL_MACHINE";
            var keyPath = KeyPathTextBox.Text;
            var valueName = ValueNameTextBox.Text;
            var detectionType = DetectionTypeCombo.SelectedIndex;
            
            string preview;
            string fullPath = $"{rootKey}\\{keyPath}";
            
            if (!string.IsNullOrWhiteSpace(valueName))
            {
                fullPath += $"\\{valueName}";
            }
            
            switch (detectionType)
            {
                case 0: // Exists
                    if (string.IsNullOrWhiteSpace(valueName))
                        preview = $"Registry key exists: {fullPath}";
                    else
                        preview = $"Registry value exists: {fullPath}";
                    break;
                case 1: // String
                    var stringOperator = GetOperatorText();
                    preview = $"String value: {fullPath} {stringOperator} '{ExpectedValueTextBox.Text}'";
                    break;
                case 2: // Integer
                    var intOperator = GetOperatorText();
                    preview = $"Integer value: {fullPath} {intOperator} {ExpectedValueTextBox.Text}";
                    break;
                case 3: // Version
                    var versionOperator = GetOperatorText();
                    preview = $"Version value: {fullPath} {versionOperator} {ExpectedValueTextBox.Text}";
                    break;
                default:
                    preview = "Invalid detection type";
                    break;
            }
            
            PreviewTextBlock.Text = preview;
        }

        private string GetOperatorText()
        {
            if (OperatorCombo.SelectedIndex == -1) return "=";
            
            return OperatorCombo.SelectedIndex switch
            {
                0 => "=",
                1 => "!=",
                2 => ">",
                3 => ">=",
                4 => "<",
                5 => "<=",
                _ => "="
            };
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(KeyPathTextBox.Text))
            {
                MessageBox.Show("Please specify a registry key path.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var rootKey = ((ComboBoxItem)RootKeyCombo.SelectedItem)?.Content?.ToString() ?? "HKEY_LOCAL_MACHINE";
            var fullKeyPath = $"{rootKey}\\{KeyPathTextBox.Text}";

            // Create the detection rule
            DetectionRule = new RegistryDetectionRule
            {
                Icon = "🗃️",
                KeyPath = fullKeyPath,
                ValueName = ValueNameTextBox.Text,
                DetectionType = (RegistryDetectionType)DetectionTypeCombo.SelectedIndex,
                ExpectedValue = ExpectedValueTextBox.Text
            };

            // Set title and description based on detection type
            switch (DetectionTypeCombo.SelectedIndex)
            {
                case 0: // Exists
                    DetectionRule.Title = "Registry Key/Value Detection";
                    if (string.IsNullOrWhiteSpace(ValueNameTextBox.Text))
                        DetectionRule.Description = $"Key exists: {fullKeyPath}";
                    else
                        DetectionRule.Description = $"Value exists: {fullKeyPath}\\{ValueNameTextBox.Text}";
                    break;
                case 1: // String
                    DetectionRule.Title = "Registry String Value Detection";
                    DetectionRule.Description = $"String {GetOperatorText()} '{ExpectedValueTextBox.Text}': {ValueNameTextBox.Text}";
                    break;
                case 2: // Integer
                    DetectionRule.Title = "Registry Integer Value Detection";
                    DetectionRule.Description = $"Integer {GetOperatorText()} {ExpectedValueTextBox.Text}: {ValueNameTextBox.Text}";
                    break;
                case 3: // Version
                    DetectionRule.Title = "Registry Version Value Detection";
                    DetectionRule.Description = $"Version {GetOperatorText()} {ExpectedValueTextBox.Text}: {ValueNameTextBox.Text}";
                    break;
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