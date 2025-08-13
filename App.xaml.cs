using System.Windows;

namespace IntunePackagingTool
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Set up global exception handling
            this.DispatcherUnhandledException += (sender, ex) =>
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Exception.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ex.Handled = true;
            };
        }
    }
}