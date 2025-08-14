using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace IntunePackagingTool
{
    public partial class MainWindow : Window
    {
        private IntuneService _intuneService = new IntuneService();
        private PSADTGenerator _psadtGenerator = new PSADTGenerator();
        private ObservableCollection<IntuneApplication> _applications = new ObservableCollection<IntuneApplication>();
        private string _currentPackagePath = "";

        public MainWindow()
        {
            InitializeComponent();
            ApplicationsList.ItemsSource = _applications;
            _ = LoadApplicationsAsync();
        }

        private async Task LoadApplicationsAsync()
        {
            try
            {
            ShowStatus("Loading applications from Intune...");
            ShowProgress(true);

            var apps = await _intuneService.GetApplicationsAsync();
        
            _applications.Clear();
            foreach (var app in apps)
            {  
            _applications.Add(app);
            }

            // Populate category filter with actual categories from loaded apps
            UpdateCategoryFilter();

            ShowStatus($"Loaded {apps.Count} applications");
            ShowProgress(false);
            }
            catch (Exception ex)
        {
        ShowStatus("Failed to load applications");
        ShowProgress(false);
        MessageBox.Show($"Error loading applications: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

        private void UpdateCategoryFilter()
        {
         var currentSelection = CategoryFilter.SelectedItem?.ToString();
    
        CategoryFilter.Items.Clear();
        CategoryFilter.Items.Add("All Categories");
    
        // Get unique categories from loaded applications
        var categories = _applications
        .Where(app => !string.IsNullOrWhiteSpace(app.Category))
        .Select(app => app.Category)
        .Distinct()
        .OrderBy(cat => cat);
    
        foreach (var category in categories)
        {
            CategoryFilter.Items.Add(category);
        }
    
        // Restore selection or default to "All Categories"
        var itemToSelect = CategoryFilter.Items.Cast<string>().FirstOrDefault(item => item == currentSelection) ?? "All Categories";
        CategoryFilter.SelectedItem = itemToSelect;
        FilterApplications();
     }

        private void ViewAppsNavButton_Click(object sender, RoutedEventArgs e)
        {
            ViewApplicationsPage.Visibility = Visibility.Visible;
            CreateApplicationPage.Visibility = Visibility.Collapsed;
            
            ViewAppsNavButton.Style = (Style)FindResource("ActiveNavButton");
            CreateAppNavButton.Style = (Style)FindResource("NavButton");
        }

        private void CreateAppNavButton_Click(object sender, RoutedEventArgs e)
        {
            ViewApplicationsPage.Visibility = Visibility.Collapsed;
            CreateApplicationPage.Visibility = Visibility.Visible;
            
            ViewAppsNavButton.Style = (Style)FindResource("NavButton");
            CreateAppNavButton.Style = (Style)FindResource("ActiveNavButton");
        }

        private void CategoryFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FilterApplications();
        }
        private void FilterApplications()
    {
    if (CategoryFilter.SelectedItem == null) return;

    var selectedCategory = CategoryFilter.SelectedItem.ToString();
    
    if (selectedCategory == "All Categories")
        {
        // Show all applications
        ApplicationsList.ItemsSource = _applications;
        }
        else
        {
        // Filter applications by selected category
        var filteredApps = _applications.Where(app => 
            !string.IsNullOrWhiteSpace(app.Category) && 
            app.Category.Equals(selectedCategory, StringComparison.OrdinalIgnoreCase))
            .ToList();
        
        ApplicationsList.ItemsSource = filteredApps;
        }
    }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadApplicationsAsync();
        }

        private async void DebugTest_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _intuneService.RunFullDebugTestAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Debug test failed: {ex.Message}", "Debug Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseSourcesButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Source Files",
                Filter = "All files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                SourcesPathTextBox.Text = string.Join(";", dialog.FileNames);
            }
        }

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateInput())
                    return;

                ShowStatus("Generating PSADT package...");
                ShowProgress(true);

                var appInfo = new ApplicationInfo
                {
                    Manufacturer = ManufacturerTextBox.Text.Trim(),
                    Name = AppNameTextBox.Text.Trim(),
                    Version = VersionTextBox.Text.Trim(),
                    SourcesPath = SourcesPathTextBox.Text.Trim(),
                    ServiceNowSRI = "SRI-" + DateTime.Now.ToString("yyyyMMdd")
                };

                _currentPackagePath = await _psadtGenerator.CreatePackageAsync(appInfo);

                PackageStatusPanel.Visibility = Visibility.Visible;
                PackageStatusText.Text = "✅ Package created successfully";
                PackagePathText.Text = _currentPackagePath;

                ShowStatus("Package generated successfully");
                ShowProgress(false);

                MessageBox.Show($"Package created successfully at:\n{_currentPackagePath}", "Success", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ShowStatus("Failed to generate package");
                ShowProgress(false);
                MessageBox.Show($"Error generating package: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenScriptButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_currentPackagePath))
                {
                    MessageBox.Show("Please generate a package first.", "No Package", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var scriptPath = Path.Combine(_currentPackagePath, "Application", "Deploy-Application.ps1");
                
                if (!File.Exists(scriptPath))
                {
                    MessageBox.Show("Script file not found.", "File Not Found", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell_ise.exe",
                    Arguments = $"\"{scriptPath}\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening script: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_currentPackagePath))
                {
                    MessageBox.Show("Please generate a package first.", "No Package", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var uploadWindow = new UploadToIntuneWindow
                {
                    Owner = this,
                    ApplicationInfo = new ApplicationInfo
                    {
                        Manufacturer = ManufacturerTextBox.Text.Trim(),
                        Name = AppNameTextBox.Text.Trim(),
                        Version = VersionTextBox.Text.Trim()
                    },
                    PackagePath = _currentPackagePath
                };

                uploadWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening upload window: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(ManufacturerTextBox.Text))
            {
                MessageBox.Show("Please enter a manufacturer name.", "Validation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                ManufacturerTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(AppNameTextBox.Text))
            {
                MessageBox.Show("Please enter an application name.", "Validation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                AppNameTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(VersionTextBox.Text))
            {
                MessageBox.Show("Please enter a version number.", "Validation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                VersionTextBox.Focus();
                return false;
            }

            return true;
        }

        private void ShowStatus(string message)
        {
            StatusText.Text = message;
        }

        private void ShowProgress(bool show)
        {
            ProgressBar.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (show)
            {
                ProgressBar.IsIndeterminate = true;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _intuneService?.Dispose();
        }
    }
}