using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using DCSBIOSBridge.Properties;
using DCS_BIOS;
using DCS_BIOS.EventArgs;
using DCS_BIOS.Interfaces;
using DCSBIOSBridge.Events;
using DCSBIOSBridge.Interfaces;
using DCSBIOSBridge.misc;
using DCSBIOSBridge.SerialPortClasses;
using DCSBIOSBridge.UserControls;
using DCSBIOSBridge.Windows;
using NLog;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Input;
using Octokit;
using System.Windows.Navigation;
using DCSBIOSBridge.Events.Args;
using Cursors = System.Windows.Input.Cursors;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;

namespace DCSBIOSBridge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : IDcsBiosConnectionListener,
        ISerialPortStatusListener, ISettingsDirtyListener, ISerialPortUserControlListener, IWindowsSerialPortEventListener, IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly object _lockObject = new();
        private const string WindowName = "DCS-BIOS Bridge ";
        private DCSBIOS _dcsBios;
        private bool _formLoaded;
        private bool _isDirty;
        private readonly SerialPortsProfileHandler _profileHandler = new();
        private readonly SerialPortService _serialPortService = new();
        private List<SerialPortUserControl> _serialPortUserControls = new();
        private readonly Dictionary<string, CancellationTokenSource> _pendingRemovedPortRescans = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TaskCompletionSource<bool>> _pendingRestoredPortOpenSignals = new(StringComparer.OrdinalIgnoreCase);
        private const int RemovedPortRescanIntervalMilliseconds = 10000;
        private const int RemovedPortRescanMaxDurationMilliseconds = 60000;
        private static readonly int[] RestoredPortOpenRetryDelaysMilliseconds = [10000, 20000, 30000];
        private const int RestoredPortOpenAttemptWaitMilliseconds = 7000;

        public MainWindow()
        {
            InitializeComponent();
            DBEventManager.AttachSerialPortStatusListener(this);
            DBEventManager.AttachWindowsSerialPortEventListener(this);
            DBEventManager.AttachSerialPortUserControlListener(this);
            DBEventManager.AttachSettingsDirtyListener(this);
            BIOSEventHandler.AttachConnectionListener(this);
        }

        #region IDisposable Support
        private bool _hasBeenCalledAlready; // To detect redundant calls

        private void Dispose(bool disposing)
        {
            if (_hasBeenCalledAlready) return;

            if (disposing)
            {
                //  dispose managed state (managed objects).
                _dcsBios?.Shutdown();
                _dcsBios?.Dispose();

                DBEventManager.BroadCastSerialPortUserControlStatus(SerialPortUserControlStatus.DoDispose);
                _serialPortService.Dispose();
                DBEventManager.DetachSerialPortStatusListener(this);
                DBEventManager.DetachWindowsSerialPortEventListener(this);
                DBEventManager.DetachSerialPortUserControlListener(this);
                DBEventManager.DetachSettingsDirtyListener(this);
                BIOSEventHandler.DetachConnectionListener(this);
                CancelAllPendingRemovedPortRescans();
                CancelAllPendingRestoredPortOpenSignals();
            }

            //  free unmanaged resources (unmanaged objects) and override a finalizer below.

            //  set large fields to null.
            _hasBeenCalledAlready = true;
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);

            //  uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }
        #endregion


        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_formLoaded) return;

                Top = Settings.Default.MainWindowTop;
                Left = Settings.Default.MainWindowLeft;
                Height = Settings.Default.MainWindowHeight;
                Width = Settings.Default.MainWindowWidth;
                
                CreateDCSBIOS();

                CheckForNewRelease();

                LoadPorts();

                if (Settings.Default.OpenAllPortsOnStartup)
                {
                    DBEventManager.BroadCastPortStatus(null, SerialPortStatus.Open);
                }

                SetShowInfoMenuItems();
                AppThemeManager.ApplyTitleBarTheme(this);

                SetWindowState();
                _formLoaded = true;
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        public void DcsBiosConnectionActive(object sender, DCSBIOSConnectionEventArgs e)
        {
            try
            {
                Dispatcher?.BeginInvoke((Action)(() => ControlSpinningWheel.RotateGear()));
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        public void OnSerialPortStatusChanged(SerialPortStatusEventArgs e)
        {
            try
            {
                switch (e.SerialPortStatus)
                {
                    case SerialPortStatus.Settings:
                        {
                            Dispatcher.Invoke(SetWindowState);
                            break;
                        }
                    case SerialPortStatus.Open:
                        break;
                    case SerialPortStatus.Close:
                        break;
                    case SerialPortStatus.Opened:
                        {
                            CompleteRestoredPortOpenSignal(e.SerialPortName, true);
                            break;
                        }
                    case SerialPortStatus.Closed:
                        break;
                    case SerialPortStatus.Added:
                    case SerialPortStatus.Hidden:
                        {
                            break;
                        }
                    case SerialPortStatus.None:
                        break;
                    case SerialPortStatus.Ok:
                        break;
                    case SerialPortStatus.Error:
                        {
                            if (!string.IsNullOrWhiteSpace(e.SerialPortName))
                            {
                                Logger.Warn($"Error status received for {e.SerialPortName}. Scheduling removed-port rescan.");
                                CompleteRestoredPortOpenSignal(e.SerialPortName, false);
                                ScheduleRemovedPortRescan(e.SerialPortName);
                            }
                            break;
                        }
                    case SerialPortStatus.Critical:
                        {
                            if (!string.IsNullOrWhiteSpace(e.SerialPortName))
                            {
                                Logger.Warn($"Critical serial port status received for {e.SerialPortName}. Scheduling removed-port rescan.");
                                CompleteRestoredPortOpenSignal(e.SerialPortName, false);
                                ScheduleRemovedPortRescan(e.SerialPortName);
                            }
                            break;
                        }
                    case SerialPortStatus.WatchDogBark:
                        break;
                    case SerialPortStatus.IOError:
                        {
                            if (!string.IsNullOrWhiteSpace(e.SerialPortName))
                            {
                                Logger.Warn($"IO error status received for {e.SerialPortName}. Scheduling removed-port rescan.");
                                CompleteRestoredPortOpenSignal(e.SerialPortName, false);
                                ScheduleRemovedPortRescan(e.SerialPortName);
                            }
                            break;
                        }
                    case SerialPortStatus.TimeOutError:
                        {
                            if (!string.IsNullOrWhiteSpace(e.SerialPortName))
                            {
                                Logger.Warn($"Timeout error status received for {e.SerialPortName}. Scheduling removed-port rescan.");
                                CompleteRestoredPortOpenSignal(e.SerialPortName, false);
                                ScheduleRemovedPortRescan(e.SerialPortName);
                            }
                            break;
                        }
                    case SerialPortStatus.BytesWritten:
                        break;
                    case SerialPortStatus.BytesRead:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(e.SerialPortStatus.ToString());
                }
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        public void OnSerialPortUserControlStatusChanged(SerialPortUserControlArgs args)
        {
            try
            {
                switch (args.Status)
                {
                    case SerialPortUserControlStatus.Created:
                        {
                            Dispatcher.Invoke(() => AddUserControlToUI(args.SerialPortUserControl));
                            break;
                        }
                    case SerialPortUserControlStatus.Hidden:
                        {
                            Dispatcher.Invoke(() => RemoveUserControlFromUI(args.SerialPortUserControl));
                            break;
                        }
                    case SerialPortUserControlStatus.Closed:
                        {
                            Dispatcher.Invoke(() => RemoveUserControlFromUI(args.SerialPortUserControl));
                            break;
                        }
                    case SerialPortUserControlStatus.Check:
                    case SerialPortUserControlStatus.DisposeDisabledPorts:
                    case SerialPortUserControlStatus.DoDispose:
                        break;
                    case SerialPortUserControlStatus.ShowInfo:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(args.Status.ToString());
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => Common.ShowErrorMessageBox(ex));
            }
        }

        public void PortsChangedEvent(object sender, PortsChangedArgs e)
        {
            try
            {
                var thread = new Thread(() => CheckComPortExistenceStatus(e.SerialPorts, e.EventType));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void LoadPorts()
        {
            if (!string.IsNullOrEmpty(Settings.Default.LastProfileFileUsed))
            {
                _profileHandler.LoadProfile(Settings.Default.LastProfileFileUsed);
                if (_profileHandler.SerialPortSettings.Count == 0)
                {
                    ListAllSerialPorts();
                }
                else
                {
                    SerialPortUserControl.LoadSerialPorts(_profileHandler.SerialPortSettings);
                }
            }
            else
            {
                ListAllSerialPorts();
            }
        }

        private void CheckComPortExistenceStatus(string[] comPorts, WindowsSerialPortEventType eventType)
        {
            try
            {
                lock (_lockObject)
                {
                    // make all SerialPortUserControl check whether their SerialPort is OK
                    // if new profile then don't send list, affects whether to remove or grey them out when removed from computer
                    DBEventManager.BroadCastSerialPortUserControlStatus(SerialPortUserControlStatus.Check, null, null, _profileHandler.IsNewProfile ? null : _profileHandler.SerialPortSettings);

                    switch (eventType)
                    {
                        case WindowsSerialPortEventType.Insertion:
                            {
                                foreach (var comPort in comPorts)
                                {
                                    CancelPendingRemovedPortRescan(comPort, "insertion event");

                                    if (Dispatcher.Invoke(() => _serialPortUserControls.Any(o => o.Name == comPort)) == false)
                                    {
                                        Dispatcher.Invoke(() =>
                                        {
                                            AddSerialPort(comPort);
                                            if (Settings.Default.OpenAllPortsOnStartup)
                                            {
                                                DBEventManager.BroadCastPortStatus(comPort, SerialPortStatus.Open);
                                            }
                                        });
                                    }
                                }
                                break;
                            }
                        case WindowsSerialPortEventType.Removal:
                        {
                            foreach (var comPort in comPorts)
                            {
                                ScheduleRemovedPortRescan(comPort);
                            }

                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ListAllSerialPorts()
        {
            var serialPorts= Common.GetSerialPortNames().ToList();
            _profileHandler.ClearHiddenPorts();

            foreach (var port in serialPorts)
            {
                if (_serialPortUserControls.Any(o => o.Name == port) == false)
                {
                    AddSerialPort(port);
                }
            }

            SetWindowState();
        }

        private void ScheduleRemovedPortRescan(string comPort)
        {
            if (!Settings.Default.WatchDogEnabled)
            {
                Logger.Info($"Watch Dog is disabled. Skipping removed-port rescan scheduling for {comPort}.");
                return;
            }

            if (string.IsNullOrWhiteSpace(comPort))
            {
                return;
            }

            CancellationTokenSource cancellationTokenSource;

            lock (_lockObject)
            {
                if (_pendingRemovedPortRescans.TryGetValue(comPort, out var existingTokenSource))
                {
                    Logger.Info($"Replacing pending removed-port rescan for {comPort}.");
                    existingTokenSource.Cancel();
                    existingTokenSource.Dispose();
                }

                cancellationTokenSource = new CancellationTokenSource();
                _pendingRemovedPortRescans[comPort] = cancellationTokenSource;
            }

            Logger.Info($"Scheduled removed-port rescan for {comPort} every {RemovedPortRescanIntervalMilliseconds} ms for up to {RemovedPortRescanMaxDurationMilliseconds} ms.");

            _ = TryRestoreRemovedPortAsync(comPort, cancellationTokenSource.Token);
        }

        private async Task TryRestoreRemovedPortAsync(string comPort, CancellationToken cancellationToken)
        {
            try
            {
                var attempts = 0;
                var startedAt = DateTime.UtcNow;

                while (true)
                {
                    await Task.Delay(RemovedPortRescanIntervalMilliseconds, cancellationToken);
                    attempts++;

                    var elapsed = DateTime.UtcNow - startedAt;
                    Logger.Info($"Running removed-port rescan attempt {attempts} for {comPort} at +{elapsed.TotalSeconds:0}s.");

                    var serialPorts = Common.GetSerialPortNames();
                    if (serialPorts.Any(o => o.Equals(comPort, StringComparison.OrdinalIgnoreCase)) == false)
                    {
                        if (elapsed.TotalMilliseconds >= RemovedPortRescanMaxDurationMilliseconds)
                        {
                            Logger.Warn($"Removed-port rescan timed out for {comPort} after {attempts} attempts (+{elapsed.TotalSeconds:0}s). Last available ports: {string.Join(", ", serialPorts)}");
                            return;
                        }

                        Logger.Warn($"Removed-port rescan did not find {comPort} on attempt {attempts}. Available ports: {string.Join(", ", serialPorts)}");
                        continue;
                    }

                    Dispatcher.Invoke(() =>
                    {
                        if (_serialPortUserControls.Any(o => o.Name.Equals(comPort, StringComparison.OrdinalIgnoreCase)))
                        {
                            Logger.Info($"Removed-port rescan found {comPort}, but UI control already exists. Skipping re-add.");
                            return;
                        }

                        AddSerialPort(comPort);
                        DBEventManager.BroadCastPortStatus(comPort, SerialPortStatus.WatchDogBark);

                        Logger.Info($"Removed-port rescan re-added {comPort} and emitted WatchDogBark.");
                    });

                    if (Settings.Default.OpenAllPortsOnStartup)
                    {
                        await TryOpenRestoredPortWithRetriesAsync(comPort, cancellationToken);
                    }

                    return;
                }
            }
            catch (OperationCanceledException)
            {
                Logger.Info($"Removed-port rescan canceled for {comPort}.");
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
            finally
            {
                lock (_lockObject)
                {
                    if (_pendingRemovedPortRescans.TryGetValue(comPort, out var tokenSource) && tokenSource.Token == cancellationToken)
                    {
                        _pendingRemovedPortRescans.Remove(comPort);
                        tokenSource.Dispose();
                    }
                }
            }
        }

        private async Task TryOpenRestoredPortWithRetriesAsync(string comPort, CancellationToken cancellationToken)
        {
            Logger.Info($"OpenAllPortsOnStartup is enabled. Starting delayed open retries for restored port {comPort}.");

            for (var i = 0; i < RestoredPortOpenRetryDelaysMilliseconds.Length; i++)
            {
                var delayMs = RestoredPortOpenRetryDelaysMilliseconds[i];
                await Task.Delay(delayMs, cancellationToken);

                if (Dispatcher.Invoke(() => _serialPortUserControls.Any(o => o.Name.Equals(comPort, StringComparison.OrdinalIgnoreCase))) == false)
                {
                    Logger.Info($"Restored port {comPort} no longer exists in UI before open attempt {i + 1}. Stopping open retries.");
                    return;
                }

                var openSignal = RegisterRestoredPortOpenSignal(comPort);

                Logger.Info($"Opening restored port {comPort} on retry attempt {i + 1} after waiting {delayMs} ms.");
                DBEventManager.BroadCastPortStatus(comPort, SerialPortStatus.Open);

                var completedTask = await Task.WhenAny(openSignal.Task, Task.Delay(RestoredPortOpenAttemptWaitMilliseconds, cancellationToken));
                if (completedTask == openSignal.Task && openSignal.Task.Result)
                {
                    Logger.Info($"Restored port {comPort} opened successfully on retry attempt {i + 1}.");
                    return;
                }

                Logger.Warn($"Restored port {comPort} did not open successfully on retry attempt {i + 1}.");
            }

            Logger.Warn($"Restored port {comPort} failed to open after {RestoredPortOpenRetryDelaysMilliseconds.Length} retry attempts.");
        }

        private TaskCompletionSource<bool> RegisterRestoredPortOpenSignal(string comPort)
        {
            lock (_lockObject)
            {
                if (_pendingRestoredPortOpenSignals.TryGetValue(comPort, out var existingSignal))
                {
                    existingSignal.TrySetCanceled();
                }

                var signal = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                _pendingRestoredPortOpenSignals[comPort] = signal;
                return signal;
            }
        }

        private void CompleteRestoredPortOpenSignal(string comPort, bool success)
        {
            if (string.IsNullOrWhiteSpace(comPort))
            {
                return;
            }

            lock (_lockObject)
            {
                if (_pendingRestoredPortOpenSignals.TryGetValue(comPort, out var signal) == false)
                {
                    return;
                }

                _pendingRestoredPortOpenSignals.Remove(comPort);
                signal.TrySetResult(success);
            }
        }

        private void CancelAllPendingRestoredPortOpenSignals()
        {
            lock (_lockObject)
            {
                foreach (var signal in _pendingRestoredPortOpenSignals.Values)
                {
                    signal.TrySetCanceled();
                }

                _pendingRestoredPortOpenSignals.Clear();
            }
        }

        private void CancelPendingRemovedPortRescan(string comPort, string reason = "unspecified")
        {
            if (string.IsNullOrWhiteSpace(comPort))
            {
                return;
            }

            lock (_lockObject)
            {
                if (_pendingRemovedPortRescans.TryGetValue(comPort, out var tokenSource) == false)
                {
                    return;
                }

                _pendingRemovedPortRescans.Remove(comPort);
                tokenSource.Cancel();
                tokenSource.Dispose();
                Logger.Info($"Canceled pending removed-port rescan for {comPort}. Reason: {reason}.");
            }
        }

        private void CancelAllPendingRemovedPortRescans()
        {
            lock (_lockObject)
            {
                var cancelledCount = _pendingRemovedPortRescans.Count;
                foreach (var tokenSource in _pendingRemovedPortRescans.Values)
                {
                    tokenSource.Cancel();
                    tokenSource.Dispose();
                }

                _pendingRemovedPortRescans.Clear();
                if (cancelledCount > 0)
                {
                    Logger.Info($"Canceled {cancelledCount} pending removed-port rescans during cleanup.");
                }
            }
        }

        private void OpenProfile()
        {
            if (!DiscardChanges()) return;

            var openFileDialog = Common.OpenFileDialog(Settings.Default.LastProfileFileUsed);

            if (openFileDialog.ShowDialog() == true)
            {
                DisposeAllUserControls();
                _profileHandler.LoadProfile(openFileDialog.FileName);
                SerialPortUserControl.LoadSerialPorts(_profileHandler.SerialPortSettings);
            }
            SetWindowState();
        }

        private void SaveAsNewProfile()
        {
            var lastDirectory = string.IsNullOrEmpty(Settings.Default.LastProfileFileUsed) ? "" : Path.GetDirectoryName(Settings.Default.LastProfileFileUsed);
            var saveFileDialog = Common.SaveProfileDialog(lastDirectory);
            if (saveFileDialog.ShowDialog() == true)
            {
                _profileHandler.SaveProfileAs(saveFileDialog.FileName);
            }
            SetWindowState();
        }

        private void SaveNewOrExistingProfile()
        {
            if (_profileHandler.IsNewProfile)
            {
                SaveAsNewProfile();
            }
            else
            {
                _profileHandler.SaveProfile();
            }
            SetWindowState();
        }

        private void SetWindowState()
        {
            ButtonImageSave.IsEnabled = _isDirty;
            MenuItemSave.IsEnabled = _isDirty && !_profileHandler.IsNewProfile;
            MenuItemSaveAs.IsEnabled = true;
            MenuItemOpen.IsEnabled = true;
            ButtonImageNotepad.IsEnabled = !_profileHandler.IsNewProfile && !_isDirty;
            SetWindowTitle();
        }

        private void SetWindowTitle()
        {
            Title = _profileHandler.IsNewProfile ? WindowName : WindowName + _profileHandler.FileName;

            if (_isDirty)
            {
                Title += " *";
            }
        }

        private void MenuItemOpenClick(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenProfile();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void MenuItemSave_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _profileHandler.SaveProfile();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void MenuItemSaveAs_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveAsNewProfile();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void MenuItemExit_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void MenuItemOptions_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var settingsWindow = new SettingsWindow
                {
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                };
                if (settingsWindow.ShowDialog() == true)
                {
                    CreateDCSBIOS();
                }
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void MenuItemAbout_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var about = new AboutWindow();
                about.ShowDialog();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void AddSerialPort(string serialPortName)
        {
            try
            {
                if (_serialPortUserControls.Count(o => o.Name == serialPortName) > 0) return;

                var existingPortSetting = _profileHandler.SerialPortSettings.FirstOrDefault(o => o.ComPort == serialPortName);
                var serialPortSetting = existingPortSetting ?? new SerialPortSetting { ComPort = serialPortName };

                var serialPortUserControl = new SerialPortUserControl(serialPortSetting, (HardwareInfoToShow)Settings.Default.ShowInfoType);
                AddUserControlToUI(serialPortUserControl);
                DBEventManager.BroadCastPortStatus(serialPortName, SerialPortStatus.Added, 0, null, serialPortUserControl.SerialPortSetting);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonNew_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isDirty && MessageBox.Show("Discard unsaved changes to current profile?", "Discard changes?", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                {
                    return;
                }

                DisposeAllUserControls();
                _profileHandler.Reset();
                ListAllSerialPorts();
                SetWindowState();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonSave_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveNewOrExistingProfile();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonSearchForSerialPorts_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                ListAllSerialPorts();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonOpen_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenProfile();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonOpenInEditor_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(_profileHandler.FileName);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private bool DiscardChanges()
        {
            if (_isDirty && MessageBox.Show("Discard unsaved changes to current profile?", "Discard changes?", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
                return false;
            }

            return true;
        }

        private void MainWindow_OnClosing(object sender, CancelEventArgs e)
        {
            try
            {
                Settings.Default.MainWindowTop = Top;
                Settings.Default.MainWindowLeft = Left;
                Settings.Default.MainWindowHeight = Height;
                Settings.Default.MainWindowWidth = Width;
                Settings.Default.Save();

                if (!DiscardChanges())
                {
                    e.Cancel = true;
                    return;
                }

                Dispose(true);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }


        private void AddUserControlToUI(SerialPortUserControl userControl)
        {
            _serialPortUserControls.Add(userControl);
            _serialPortUserControls = _serialPortUserControls.OrderBy(o => o.Name).ToList();
            ItemsControlPorts.ItemsSource = null;
            ItemsControlPorts.ItemsSource = _serialPortUserControls;
        }

        private void RemoveUserControlFromUI(SerialPortUserControl userControl)
        {
            _serialPortUserControls.Remove(userControl);
            ItemsControlPorts.ItemsSource = null;
            ItemsControlPorts.ItemsSource = _serialPortUserControls;
        }

        private void DisposeAllUserControls()
        {
            try
            {
                DBEventManager.BroadCastSerialPortUserControlStatus(SerialPortUserControlStatus.DoDispose);
                _serialPortUserControls.Clear();
                ItemsControlPorts.ItemsSource = null;
                ItemsControlPorts.ItemsSource = _serialPortUserControls;
                SetWindowState();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void CreateDCSBIOS()
        {
            _dcsBios?.Shutdown();
            _dcsBios?.Dispose();
            _dcsBios = new DCSBIOS(Settings.Default.DCSBiosIPFrom,
                Settings.Default.DCSBiosIPTo,
                int.Parse(Settings.Default.DCSBiosPortFrom),
                int.Parse(Settings.Default.DCSBiosPortTo),
                DcsBiosNotificationMode.PassThrough);

            if (!_dcsBios.HasLastException())
            {
                ControlSpinningWheel.RotateGear(2000);
            }

            _dcsBios.DelayBetweenCommands = Settings.Default.DelayBetweenCommands;
        }

        private void MenuItemLogFile_OnClick(object sender, RoutedEventArgs e)
        {
            Common.TryOpenLogFileWithTarget("logfile");
        }

        public void OnSettingsDirty(SettingsDirtyEventArgs args)
        {
            try
            {
                _isDirty = args.IsDirty;
                Dispatcher.Invoke(SetWindowState);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonClosePorts_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                DBEventManager.BroadCastPortStatus(null, SerialPortStatus.Close);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void ButtonOpenPorts_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                DBEventManager.BroadCastPortStatus(null, SerialPortStatus.Open);
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private async void CheckForNewRelease()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            if (string.IsNullOrEmpty(fileVersionInfo.FileVersion)) return;

            var thisVersion = new Version(fileVersionInfo.FileVersion);

            try
            {
                var dateTime = Settings.Default.LastGitHubCheck;

                var client = new GitHubClient(new ProductHeaderValue("DCSBIOSBridge"));
                var timeSpan = DateTime.Now - dateTime;
                if (timeSpan.Days > 1)
                {
                    Settings.Default.LastGitHubCheck = DateTime.Now;
                    Settings.Default.Save();
                    var lastRelease = await client.Repository.Release.GetLatest("DCS-Skunkworks", "DCSBIOSBridge");
                    var githubVersion = new Version(lastRelease.TagName.Replace("v", ""));
                    if (githubVersion.CompareTo(thisVersion) > 0)
                    {
                        Dispatcher?.Invoke(() =>
                        {
                            LabelVersionInformation.Visibility = Visibility.Hidden;
                            LabelDownloadNewVersion.Visibility = Visibility.Visible;
                        });
                    }
                    else
                    {
                        Dispatcher?.Invoke(() =>
                        {
                            LabelVersionInformation.Text = "v." + fileVersionInfo.FileVersion;
                            LabelVersionInformation.Visibility = Visibility.Visible;
                        });
                    }
                }
                else
                {
                    Dispatcher?.Invoke(() =>
                    {
                        LabelVersionInformation.Text = "v." + fileVersionInfo.FileVersion;
                        LabelVersionInformation.Visibility = Visibility.Visible;
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error checking for newer releases.");
                LabelVersionInformation.Text = "v." + fileVersionInfo.FileVersion;
            }
        }

        private void Hyperlink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                });
                e.Handled = true;
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void UIElement_OnMouseEnterCursorArrow(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }

        private void UIElement_OnMouseLeaveCursorArrow(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
        }

        private void MenuItemRemoveDisabledPorts_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MessageBox.Show(this, "Remove disabled ports from the configuration?", "Remove ports", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    DBEventManager.BroadCastSerialPortUserControlStatus(SerialPortUserControlStatus.DisposeDisabledPorts);
                }
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }

        private void SetShowInfoMenuItems()
        {
            foreach (var item in MenuItemShow.Items)
            {
                if (item is not MenuItem menuItem || menuItem.Tag == null) continue;
                menuItem.IsChecked = int.Parse(menuItem.Tag.ToString() ?? "0") == Settings.Default.ShowInfoType;
            }

            DBEventManager.BroadCastSerialPortUserControlStatus(SerialPortUserControlStatus.ShowInfo, null, null, null, (HardwareInfoToShow)Settings.Default.ShowInfoType);
        }

        private void MenuItemShowInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var menuItem = (MenuItem)sender;

                Settings.Default.ShowInfoType = int.Parse(menuItem.Tag.ToString() ?? "0");
                Settings.Default.Save();

                SetShowInfoMenuItems();
            }
            catch (Exception ex)
            {
                Common.ShowErrorMessageBox(ex);
            }
        }
    }
}