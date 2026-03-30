using System.Diagnostics;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;
using DCS_BIOS;
using DCSBIOSBridge.Properties;

namespace DCSBIOSBridge.Windows
{
    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow
    {
        private string IpAddressFromDCSBIOS { get; set; }
        private string PortFromDCSBIOS { get; set; }
        private string IpAddressToDCSBIOS { get; set; }
        private string PortToDCSBIOS { get; set; }
        private bool DCSBIOSChanged { get; set; }

        private bool _isLoaded;

        public SettingsWindow()
        {
            InitializeComponent();
        }

        private void SetFormState() {}

        private void SettingsWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isLoaded)
                {
                    return;
                }

                Mouse.OverrideCursor = Cursors.Arrow;
                ButtonOk.IsEnabled = false;
                LoadSettings();
                SetEventsHandlers();
                SetFormState();
                _isLoaded = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + Environment.NewLine + exception.StackTrace);
            }
        }

        private void SetEventsHandlers()
        {
            TextBoxDCSBIOSFromIP.TextChanged += DcsBiosDirty;
            TextBoxDCSBIOSToIP.TextChanged += DcsBiosDirty;
            TextBoxDCSBIOSFromPort.TextChanged += DcsBiosDirty;
            TextBoxDCSBIOSToPort.TextChanged += DcsBiosDirty;
            TextBoxDCSBIOSCommandDelay.TextChanged += DcsBiosDirty;
            CheckBoxOpenAllPortsOnStartup.Checked += BridgeSettingsDirty;
            CheckBoxOpenAllPortsOnStartup.Unchecked += BridgeSettingsDirty;
            TextBoxWatchDogNoReadTimeoutSeconds.TextChanged += DcsBiosDirty;
            TextBoxWatchDogRecentWriteWindowSeconds.TextChanged += DcsBiosDirty;
            TextBoxWatchDogCooldownSeconds.TextChanged += DcsBiosDirty;
            TextBoxWatchDogReopenDelayMilliseconds.TextChanged += DcsBiosDirty;
        }

        private void LoadSettings()
        {
            TextBoxDCSBIOSFromIP.Text = Settings.Default.DCSBiosIPFrom;
            TextBoxDCSBIOSToIP.Text = Settings.Default.DCSBiosIPTo;
            TextBoxDCSBIOSFromPort.Text = Settings.Default.DCSBiosPortFrom;
            TextBoxDCSBIOSToPort.Text = Settings.Default.DCSBiosPortTo;
            TextBoxDCSBIOSCommandDelay.Text = Settings.Default.DelayBetweenCommands.ToString();
            CheckBoxOpenAllPortsOnStartup.IsChecked = Settings.Default.OpenAllPortsOnStartup;
            TextBoxWatchDogNoReadTimeoutSeconds.Text = Settings.Default.WatchDogNoReadTimeoutSeconds.ToString();
            TextBoxWatchDogRecentWriteWindowSeconds.Text = Settings.Default.WatchDogRecentWriteWindowSeconds.ToString();
            TextBoxWatchDogCooldownSeconds.Text = Settings.Default.WatchDogCooldownSeconds.ToString();
            TextBoxWatchDogReopenDelayMilliseconds.Text = Settings.Default.WatchDogReopenDelayMilliseconds.ToString();
        }

        private void ButtonCancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ButtonOk_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckValuesDCSBIOS();

                if (DCSBIOSChanged)
                {
                    Settings.Default.DCSBiosIPFrom = IpAddressFromDCSBIOS;
                    Settings.Default.DCSBiosPortFrom = PortFromDCSBIOS;
                    Settings.Default.DCSBiosIPTo = IpAddressToDCSBIOS;
                    Settings.Default.DCSBiosPortTo = PortToDCSBIOS;
                    Settings.Default.DelayBetweenCommands = Convert.ToInt32(TextBoxDCSBIOSCommandDelay.Text);
                    Settings.Default.OpenAllPortsOnStartup = CheckBoxOpenAllPortsOnStartup.IsChecked == true;
                    Settings.Default.WatchDogNoReadTimeoutSeconds = Convert.ToInt32(TextBoxWatchDogNoReadTimeoutSeconds.Text);
                    Settings.Default.WatchDogRecentWriteWindowSeconds = Convert.ToInt32(TextBoxWatchDogRecentWriteWindowSeconds.Text);
                    Settings.Default.WatchDogCooldownSeconds = Convert.ToInt32(TextBoxWatchDogCooldownSeconds.Text);
                    Settings.Default.WatchDogReopenDelayMilliseconds = Convert.ToInt32(TextBoxWatchDogReopenDelayMilliseconds.Text);
                    Settings.Default.Save();

                    DCSBIOS.GetInstance().DelayBetweenCommands = Convert.ToInt32(TextBoxDCSBIOSCommandDelay.Text);
                }
                
                DialogResult = true;
                Close();
            }
            catch (Exception exception)
            {
                MessageBox.Show($"{exception.Message}{Environment.NewLine}{exception.StackTrace}");
            }
        }

        private void CheckValuesDCSBIOS()
        {
            try
            {
                if (string.IsNullOrEmpty(TextBoxDCSBIOSFromIP.Text))
                {
                    throw new Exception("DCS-BIOS IP address from cannot be empty");
                }
                if (string.IsNullOrEmpty(TextBoxDCSBIOSToIP.Text))
                {
                    throw new Exception("DCS-BIOS IP address to cannot be empty");
                }
                if (string.IsNullOrEmpty(TextBoxDCSBIOSFromPort.Text))
                {
                    throw new Exception("DCS-BIOS Port from cannot be empty");
                }
                if (string.IsNullOrEmpty(TextBoxDCSBIOSToPort.Text))
                {
                    throw new Exception("DCS-BIOS Port to cannot be empty");
                }
                try
                {
                    if (!IPAddress.TryParse(TextBoxDCSBIOSFromIP.Text, out _))
                    {
                        throw new Exception();
                    }
                    IpAddressFromDCSBIOS = TextBoxDCSBIOSFromIP.Text;
                }
                catch (Exception ex)
                {
                    throw new Exception($"DCS-BIOS Error while checking IP from : {ex.Message}");
                }
                try
                {
                    if (!IPAddress.TryParse(TextBoxDCSBIOSToIP.Text, out _))
                    {
                        throw new Exception();
                    }
                    IpAddressToDCSBIOS = TextBoxDCSBIOSToIP.Text;
                }
                catch (Exception ex)
                {
                    throw new Exception($"DCS-BIOS Error while checking IP to : {ex.Message}");
                }
                try
                {
                    _ = Convert.ToInt32(TextBoxDCSBIOSFromPort.Text);
                    PortFromDCSBIOS = TextBoxDCSBIOSFromPort.Text;
                }
                catch (Exception ex)
                {
                    throw new Exception($"DCS-BIOS Error while Port from : {ex.Message}");
                }
                try
                {
                    _ = Convert.ToInt32(TextBoxDCSBIOSFromPort.Text);
                    PortToDCSBIOS = TextBoxDCSBIOSToPort.Text;
                }
                catch (Exception ex)
                {
                    throw new Exception($"DCS-BIOS Error while Port to : {ex.Message}");
                }
                try
                {
                    var delayValue = Convert.ToInt32(TextBoxDCSBIOSCommandDelay.Text);
                    if (delayValue == 0)
                    {
                        throw new Exception("Delay cannot be zero.");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"DCS-BIOS Delay between DCS-BIOS commands : {ex.Message}");
                }

                ValidateGreaterThanZero(TextBoxWatchDogNoReadTimeoutSeconds.Text, "Watch Dog no-read timeout seconds");
                ValidateGreaterThanZero(TextBoxWatchDogRecentWriteWindowSeconds.Text, "Watch Dog recent-write window seconds");
                ValidateGreaterThanZero(TextBoxWatchDogCooldownSeconds.Text, "Watch Dog cooldown seconds");
                ValidateGreaterThanZero(TextBoxWatchDogReopenDelayMilliseconds.Text, "Watch Dog reopen delay milliseconds");
            }
            catch (Exception ex)
            {
                throw new Exception($"DCS-BIOS Error checking values : {Environment.NewLine}{ex.Message}");
            }
        }

        private static void ValidateGreaterThanZero(string value, string fieldName)
        {
            try
            {
                var parsed = Convert.ToInt32(value);
                if (parsed <= 0)
                {
                    throw new Exception($"{fieldName} must be greater than zero.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"{fieldName} : {ex.Message}");
            }
        }
        
        private void DcsBiosDirty(object sender, TextChangedEventArgs e)
        {
            DCSBIOSChanged = true;
            ButtonOk.IsEnabled = true;
        }

        private void BridgeSettingsDirty(object sender, RoutedEventArgs e)
        {
            DCSBIOSChanged = true;
            ButtonOk.IsEnabled = true;
        }
        
        private void SettingsWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (ButtonOk.IsEnabled || e.Key != Key.Escape) return;

            DialogResult = false;
            e.Handled = true;
            Close();
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
        
    }
}
