using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;
using DCSBIOSBridge.misc;

namespace DCSBIOSBridge.Windows
{
    /// <summary>
    /// Interaction logic for AboutWindow.xaml
    /// </summary>
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            AppThemeManager.ApplyTitleBarTheme(this);
        }

        private void HyperlinkRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
            e.Handled = true;
        }

        private void ButtonOk_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
