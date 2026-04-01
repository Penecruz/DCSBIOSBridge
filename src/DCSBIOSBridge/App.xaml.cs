using System.Windows;
using DCSBIOSBridge.misc;

namespace DCSBIOSBridge
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_OnStartup(object sender, StartupEventArgs e)
        {
            AppThemeManager.Initialize();
        }

        private void Application_OnExit(object sender, ExitEventArgs e)
        {
            AppThemeManager.Shutdown();
        }
    }

}
