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
        }

        private void DetectionMethodCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ValueComparisonPanel == null) return;

            var selectedItem = (ComboBoxItem)DetectionMethodCombo.SelectedItem;
            var content = selectedItem?.Content?.ToString() ?? "";

            // Show value comparison panel for comparison methods
            if (content == "String comparison" || content == "Integer comparison" || content == "Version comparison")
            {
                ValueComparisonPanel.Visibility = Visibility.Visible;
            }
            else
            {
                ValueComparisonPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (string.IsNullOrWhiteSpace(KeyPathCombo.Text))
            {
                MessageBox.Show("Please enter a registry key path.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedMethod = ((ComboBoxItem)DetectionMethodCombo.SelectedItem)?.Content?.ToString() ?? "";
            
            // Validate value name for certain detection methods
            if ((selectedMethod == "Value exists" || selectedMethod == "String comparison" || 
                 selectedMethod == "Integer comparison" || selectedMethod == "Version comparison") &&
                string.IsNullOrWhiteSpace(ValueNameTextBox.Text))
            {
                MessageBox.Show("Please enter a value name for this detection method.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Validate comparison value for comparison methods
            if ((selectedMethod == "String comparison" || selectedMethod == "Integer comparison" || selectedMethod == "Version comparison") &&
                string.IsNullOrWhiteSpace(ComparisonValueTextBox.Text))
            {
                MessageBox.Show("Please enter a comparison value.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Create the detection rule
            DetectionRule = new DetectionRule
            {
                Type = DetectionRuleType.Registry,
                RegistryKey = KeyPathCombo.Text.Trim(),
                RegistryValueName = ValueNameTextBox.Text.Trim()
            };

            // Set detection method specific properties
            switch (selectedMethod)
            {
                case "Key exists":
                    // No additional properties needed
                    break;
                case "Value exists":
                    // No additional properties needed
                    break;
                case "String comparison":
                case "Integer comparison":
                case "Version comparison":
                    DetectionRule.ExpectedValue = ComparisonValueTextBox.Text.Trim();
                    var operatorItem = (ComboBoxItem)ComparisonOperatorCombo.SelectedItem;
                    DetectionRule.Operator = ConvertOperatorToIntune(operatorItem?.Content?.ToString() ?? "Equals");
                    break;
            }

            DialogResult = true;
            Close();
        }

        private string ConvertOperatorToIntune(string displayOperator)
        {
            return displayOperator switch
            {
                "Equals" => "equal",
                "Not equals" => "notEqual",
                "Greater than" => "greaterThan",
                "Greater than or equal" => "greaterThanOrEqual",
                "Less than" => "lessThan",
                "Less than or equal" => "lessThanOrEqual",
                _ => "equal"
            };
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}