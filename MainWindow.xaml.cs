using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.IO;
using System.Linq;

namespace IntunePackagingTool
{
    public partial class MainWindow : Window
    {
        private IntuneService? _intuneService;
        private PSADTGenerator? _psadtGenerator;
        private List<IntuneApplication>? _allApplications;

        public MainWindow()
        {
            InitializeComponent();
            InitializeServices();
            
            // Set default page and active navigation button
            ShowPage("ViewApplications");
            SetActiveNavButton(ViewAppsNavButton);
            StatusText.Text = "Ready to view applications";
        }

        private void InitializeServices()
        {
            _intuneService = new IntuneService();
            _psadtGenerator = new PSADTGenerator();
        }

        #region Navigation Methods

        private void ViewAppsNavButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPage("ViewApplications");
            SetActiveNavButton(ViewAppsNavButton);
            StatusText.Text = "Viewing applications from Intune";
            
            // Auto-load data if not already loaded
            if (_allApplications == null)
            {
                LoadIntuneData();
            }
        }

        private void CreateAppNavButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPage("CreateApplication");
            SetActiveNavButton(CreateAppNavButton);
            StatusText.Text = "Ready to create new application package";
        }

        private void ShowPage(string pageName)
        {
            // Hide all pages first
            ViewApplicationsPage.Visibility = Visibility.Collapsed;
            CreateApplicationPage.Visibility = Visibility.Collapsed;

            // Show the requested page
            switch (pageName)
            {
                case "ViewApplications":
                    ViewApplicationsPage.Visibility = Visibility.Visible;
                    break;
                case "CreateApplication":
                    CreateApplicationPage.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void SetActiveNavButton(Button activeButton)
        {
            // Reset all navigation buttons to default style
            ViewAppsNavButton.Style = (Style)FindResource("NavButton");
            CreateAppNavButton.Style = (Style)FindResource("NavButton");

            // Set the active button style
            activeButton.Style = (Style)FindResource("ActiveNavButton");
        }

        #endregion

        #region Intune Data Methods

        private async void LoadIntuneData()
        {
            if (_intuneService == null)
            {
                StatusText.Text = "Intune service not initialized.";
                return;
            }

            _intuneService.EnableDebug(false);

            try
            {
                StatusText.Text = "Connecting to Microsoft Graph API...";
                ProgressBar.Visibility = Visibility.Visible;
                ProgressBar.IsIndeterminate = true;

                // First test the connection
                StatusText.Text = "Authenticating...";
                await _intuneService.GetAccessTokenAsync();

                StatusText.Text = "Loading Win32 applications with categories from Intune... (this may take a moment)";
                var apps = await _intuneService.GetApplicationsAsync();
                
                if (apps == null || apps.Count == 0)
                {
                    StatusText.Text = "No Win32 applications found";
                    MessageBox.Show("No Win32 applications were found in Intune. Please check your permissions or verify that you have Win32 LOB apps deployed.", "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // Store all applications for filtering
                _allApplications = apps;
                
                // Build dynamic category list from actual Intune categories
                var categories = new List<string> { "All Categories" };
                categories.AddRange(_allApplications
                    .SelectMany(app => app.Category.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    .Select(c => c.Trim())
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct()
                    .OrderBy(c => c));

                CategoryFilter.ItemsSource = categories;
                CategoryFilter.SelectedIndex = 0; // Default to "All Categories"
                
                // Apply current filter
                ApplyFilter();

                StatusText.Text = $"Loaded {apps.Count} Win32 applications with real-time categories from Intune successfully";
            }
            catch (Exception ex)
            {
                StatusText.Text = "Failed to load Intune data";
                
                string errorMessage = ex.Message;
                if (ex.Message.Contains("timeout") || ex.Message.Contains("ServiceUnavailable"))
                {
                    errorMessage = "The request timed out or the service is temporarily unavailable.\n\n" +
                                 "This can happen when Intune has many applications or during peak usage times.\n\n" +
                                 "Please try again in a few minutes.";
                }
                else if (ex.Message.Contains("Unauthorized") || ex.Message.Contains("Forbidden"))
                {
                    errorMessage = "Authentication failed or insufficient permissions.\n\n" +
                                 "Please check that:\n" +
                                 "• Your certificate is properly installed\n" +
                                 "• Your app registration has the correct permissions\n" +
                                 "• The certificate hasn't expired";
                }
                
                MessageBox.Show($"Error loading Win32 applications:\n\n{errorMessage}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ProgressBar.Visibility = Visibility.Collapsed;
                ProgressBar.IsIndeterminate = false;
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadIntuneData(); // Refresh with live data
        }

        private async void DebugTest_Click(object sender, RoutedEventArgs e)
        {
            if (_intuneService == null)
            {
                MessageBox.Show("Intune service not initialized.");
                return;
            }

            _intuneService.EnableDebug(true); // Enable debug messages
            await _intuneService.RunFullDebugTestAsync();
        }

        private void CategoryFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_allApplications == null || CategoryFilter.SelectedItem == null)
                return;

            // Since we're using ItemsSource with strings, SelectedItem is a string, not ComboBoxItem
            var selectedCategory = CategoryFilter.SelectedItem.ToString();
            
            List<IntuneApplication> filteredApps;
            
            if (selectedCategory == "All Categories")
            {
                filteredApps = _allApplications;
            }
            else
            {
                filteredApps = _allApplications.Where(app =>
                    app.Category.Split(',')
                        .Any(c => c.Trim().Equals(selectedCategory, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }
            
            ApplicationsList.ItemsSource = filteredApps;
            
            // Update status to show filtered count
            if (selectedCategory == "All Categories")
            {
                StatusText.Text = $"Showing all {filteredApps.Count} Win32 applications";
            }
            else
            {
                StatusText.Text = $"Showing {filteredApps.Count} Win32 applications in '{selectedCategory}' category";
            }
        }

        #endregion

        #region Package Creation Methods

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ManufacturerTextBox.Text) ||
                    string.IsNullOrWhiteSpace(AppNameTextBox.Text) ||
                    string.IsNullOrWhiteSpace(VersionTextBox.Text))
                {
                    MessageBox.Show("Please fill in Manufacturer, Application Name, and Version.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                StatusText.Text = "Creating application package...";
                ProgressBar.Visibility = Visibility.Visible;
                ProgressBar.IsIndeterminate = true;

                var appInfo = new ApplicationInfo
                {
                    Manufacturer = ManufacturerTextBox.Text.Trim(),
                    Name = AppNameTextBox.Text.Trim(),
                    Version = VersionTextBox.Text.Trim(),
                    InstallContext = "System", // Default to System for now
                    SourcesPath = SourcesPathTextBox.Text?.Trim() ?? "",
                    ServiceNowSRI = "" // Removed from UI for now
                };

                StatusText.Text = "Checking network paths...";
                if (!Directory.Exists(@"\\nbb.local\sys\SCCMData"))
                {
                    throw new DirectoryNotFoundException("Cannot access network share \\\\nbb.local\\sys\\SCCMData. Please check your network connection and permissions.");
                }

                StatusText.Text = "Creating package structure...";
                var packagePath = await _psadtGenerator!.CreatePackageAsync(appInfo);

                StatusText.Text = "Package created successfully!";
                ProgressBar.Visibility = Visibility.Collapsed;

                // Show package status
                PackageStatusPanel.Visibility = Visibility.Visible;
                PackageStatusText.Text = "✅ Package created successfully!";
                PackagePathText.Text = $"Location: {packagePath}";

                var result = MessageBox.Show($"Application package created successfully!\n\nLocation: {packagePath}\n\nWould you like to open the folder?",
                    "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = packagePath,
                        UseShellExecute = true
                    });
                }
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("already exists"))
            {
                ProgressBar.Visibility = Visibility.Collapsed;
                StatusText.Text = "Version already exists";
                MessageBox.Show($"Error: {ex.Message}\n\nPlease use a different version number.", "Version Exists", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (DirectoryNotFoundException ex)
            {
                ProgressBar.Visibility = Visibility.Collapsed;
                StatusText.Text = "Network error";
                MessageBox.Show($"Network Error: {ex.Message}", "Network Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                ProgressBar.Visibility = Visibility.Collapsed;
                StatusText.Text = "Error creating package";
                MessageBox.Show($"Error creating package: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenScriptButton_Click(object sender, RoutedEventArgs e)
        {
            string scriptPath = "";
            
            try
            {
                if (string.IsNullOrWhiteSpace(ManufacturerTextBox.Text) || string.IsNullOrWhiteSpace(AppNameTextBox.Text))
                {
                    MessageBox.Show("Please create or select an application first.", "Validation Error", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                scriptPath = _psadtGenerator!.GetScriptPath(ManufacturerTextBox.Text, AppNameTextBox.Text, VersionTextBox.Text ?? "1.0.0");
                
                // Check if the script file exists
                if (!File.Exists(scriptPath))
                {
                    var result = MessageBox.Show($"Script file not found at:\n{scriptPath}\n\nThis usually means the package hasn't been created yet. Would you like to create the package first?", 
                        "Script Not Found", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        
                    if (result == MessageBoxResult.Yes)
                    {
                        // Trigger the create package functionality
                        GenerateButton_Click(sender, e);
                        return;
                    }
                    else
                    {
                        return;
                    }
                }

                // Open the script in PowerShell ISE
                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "powershell_ise.exe",
                    Arguments = $"\"{scriptPath}\"",
                    UseShellExecute = true,
                    ErrorDialog = true
                };

                var process = System.Diagnostics.Process.Start(processStartInfo);
                
                if (process != null)
                {
                    StatusText.Text = $"Opened Deploy-Application.ps1 in PowerShell ISE";
                }
                else
                {
                    throw new Exception("Failed to start PowerShell ISE");
                }
            }
            catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 2)
            {
                // PowerShell ISE not found - try regular PowerShell with -ise parameter
                try
                {
                    var fallbackProcessStartInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoExit -Command \"ise '{scriptPath.Replace("'", "''")}' \"",
                        UseShellExecute = true,
                        ErrorDialog = true
                    };
                    
                    System.Diagnostics.Process.Start(fallbackProcessStartInfo);
                    StatusText.Text = $"Opened Deploy-Application.ps1 in PowerShell (ISE not available)";
                }
                catch
                {
                    MessageBox.Show("PowerShell ISE is not available on this system.\n\nYou can manually open the script at:\n" + scriptPath, 
                        "PowerShell ISE Not Found", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error opening script";
                MessageBox.Show($"Error opening script in PowerShell ISE:\n{ex.Message}\n\nScript location:\n{scriptPath}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseSourcesButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Use OpenFileDialog for file selection
                var fileDialog = new OpenFileDialog();
                
                // Set dialog properties
                fileDialog.Title = "Select application source files";
                fileDialog.Filter = "All Files (*.*)|*.*|Executable Files (*.exe;*.msi)|*.exe;*.msi|Setup Files (*.exe;*.msi;*.zip)|*.exe;*.msi;*.zip";
                fileDialog.FilterIndex = 1;
                fileDialog.Multiselect = true; // Allow multiple file selection
                fileDialog.CheckFileExists = true;
                fileDialog.CheckPathExists = true;
                
                // Set initial directory if SourcesPathTextBox already has a path
                var currentPath = SourcesPathTextBox.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(currentPath))
                {
                    // If it's a file path, get the directory
                    if (File.Exists(currentPath))
                    {
                        fileDialog.InitialDirectory = Path.GetDirectoryName(currentPath);
                    }
                    // If it contains multiple paths (semicolon separated), use the first one's directory
                    else if (currentPath.Contains(";"))
                    {
                        var firstPath = currentPath.Split(';')[0].Trim();
                        if (File.Exists(firstPath))
                        {
                            fileDialog.InitialDirectory = Path.GetDirectoryName(firstPath);
                        }
                    }
                    // If it's a directory path, use it
                    else if (Directory.Exists(currentPath))
                    {
                        fileDialog.InitialDirectory = currentPath;
                    }
                }
                
                if (fileDialog.InitialDirectory == null)
                {
                    // Default to Desktop if no valid path found
                    fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }

                // Show the dialog
                var result = fileDialog.ShowDialog();
                
                if (result == true && fileDialog.FileNames?.Length > 0)
                {
                    // Join multiple selected files with semicolon separator
                    var selectedFiles = string.Join(";", fileDialog.FileNames);
                    
                    // Update the textbox with the selected file paths
                    SourcesPathTextBox.Text = selectedFiles;
                    
                    // Update status to show what was selected
                    var fileCount = fileDialog.FileNames.Length;
                    StatusText.Text = $"Selected {fileCount} source file{(fileCount == 1 ? "" : "s")}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting files: {ex.Message}", "Browse Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Error selecting source files";
            }
        }

        private async void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Validate that we have the required information
                if (string.IsNullOrWhiteSpace(ManufacturerTextBox.Text) || string.IsNullOrWhiteSpace(AppNameTextBox.Text))
                {
                    MessageBox.Show("Please fill in Manufacturer and Application Name first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Check if package has been created
                var appInfo = new ApplicationInfo
                {
                    Manufacturer = ManufacturerTextBox.Text.Trim(),
                    Name = AppNameTextBox.Text.Trim(),
                    Version = VersionTextBox.Text?.Trim() ?? "1.0.0",
                    InstallContext = "System"
                };

                // Get the expected package path
                var cleanManufacturer = appInfo.Manufacturer.Replace(" ", "_");
                var cleanAppName = appInfo.Name.Replace(" ", "_");
                var packagePath = Path.Combine(@"\\nbb.local\sys\SCCMData\IntuneApplications", $"{cleanManufacturer}_{cleanAppName}", appInfo.Version);

                // Check if package exists
                if (!Directory.Exists(packagePath))
                {
                    var result = MessageBox.Show("The package hasn't been created yet. Would you like to create it first?", 
                        "Package Not Found", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        GenerateButton_Click(sender, e);
                        return;
                    }
                    else
                    {
                        return;
                    }
                }

                // Open the Upload to Intune window
                var uploadWindow = new UploadToIntuneWindow(appInfo, packagePath, _intuneService!);
                uploadWindow.Owner = this;
                
                var uploadResult = uploadWindow.ShowDialog();
                
                if (uploadResult == true)
                {
                    StatusText.Text = "Application uploaded to Intune successfully";
                    MessageBox.Show("Application has been successfully uploaded to Microsoft Intune!", 
                        "Upload Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error during upload process";
                MessageBox.Show($"Error during upload process: {ex.Message}", "Upload Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion
    }
}