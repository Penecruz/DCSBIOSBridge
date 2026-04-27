using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using DCS_BIOS.EventArgs;
using DCS_BIOS.Interfaces;
using DCSBIOSBridge.Events;
using DCSBIOSBridge.Events.Args;
using DCSBIOSBridge.Interfaces;
using DCSBIOSBridge.misc;
using DCSBIOSBridge.Properties;
using Microsoft.Win32;
using NLog;

namespace DCSBIOSBridge.SerialPortClasses
{
    public enum SerialPortStatus
    {
        Opened,
        Closed,
        Open,
        Close,
        Added,
        Hidden,
        None,
        Ok,
        IOError,
        TimeOutError,
        Error,
        Critical,
        WatchDogBark,
        BytesWritten,
        BytesRead,
        Settings
    }

    public enum HardwareInfoToShow
    {
        Name,
        VIDPID,
        CustomNameVidPidAndComPort
    }

    public class SerialPortShell : IDcsBiosBulkDataListener, ISerialDataListener, IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        //mode COM%COMPORT% BAUD=500000 PARITY=N DATA=8 STOP=1 TO=off DTR=on

        private SafeSerialPort _safeSerialPort;
        private SerialReceiver _serialReceiver;
        private readonly ConcurrentQueue<byte[]> _serialDataQueue = new();
        private AutoResetEvent _serialDataWaitingForWriteResetEvent = new(false);
        private bool _shutdown;
        private bool _portShouldBeOpen;
        private DateTime _lastSerialReadActivityUtc = DateTime.MinValue;
        private DateTime _lastSerialWriteActivityUtc = DateTime.MinValue;
        private DateTime _lastWatchDogBarkUtc = DateTime.MinValue;
        private bool _readChannelObserved;
        private int _watchDogBarkCount;
        private int _queuedMessageCount;
        private long _queuedByteCount;
        private DateTime _lastQueueDiagnosticsLogUtc = DateTime.MinValue;
        private static readonly TimeSpan QueueDiagnosticsLogInterval = TimeSpan.FromSeconds(10);
        private const int QueueMessageWarningThreshold = 100;
        private const long QueueBytesWarningThreshold = 1 * 1024 * 1024;

        public SerialPortSetting SerialPortSetting { get; set; }

        public SerialPortShell(SerialPortSetting serialPortSetting)
        {
            Debug.WriteLine($"Creating shell for {serialPortSetting.ComPort}");
            SerialPortSetting = serialPortSetting;
            GetFriendlyName();

            var thread = new Thread(CheckPortOpen);
            thread.Start();
            
            BIOSEventHandler.AttachBulkDataListener(this);
            DBEventManager.AttachDataReceivedListener(this);
        }

        #region IDisposable Support
        private bool _hasBeenCalledAlready; // To detect redundant calls

        private void Dispose(bool disposing)
        {
            if (_hasBeenCalledAlready) return;

            if (disposing)
            {
                _shutdown = true;
                _portShouldBeOpen = false;

                Debug.WriteLine($"Disposing shell for {SerialPortSetting.ComPort}");
                
                BIOSEventHandler.DetachBulkDataListener(this);
                DBEventManager.DetachDataReceivedListener(this);
                //  dispose managed state (managed objects).
                _serialDataWaitingForWriteResetEvent?.Set();
                _serialDataWaitingForWriteResetEvent?.Close();
                _serialDataWaitingForWriteResetEvent?.Dispose();
                _serialDataWaitingForWriteResetEvent = null;

                _serialReceiver?.Dispose();
                _serialReceiver = null;

                _safeSerialPort?.Close();
                _safeSerialPort?.Dispose();
                _safeSerialPort = null;
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

        public void Open(bool showErrorPopup = true)
        {
            if (_safeSerialPort != null && _safeSerialPort.IsOpen) return;

            Logger.Info($"Creating and opening serial port {SerialPortSetting.ComPort}");
 
            _serialReceiver?.Dispose();
            _serialReceiver = null;

            _safeSerialPort = new SafeSerialPort();
            _serialReceiver = new SerialReceiver
            {
                SerialPort = _safeSerialPort
            };
            _safeSerialPort.DataReceived += _serialReceiver.ReceiveTextOverSerial;
            _safeSerialPort.ErrorReceived += _serialReceiver.SerialPortError;

            ApplyPortConfig();
            var openSucceeded = false;
            try
            {
                GetFriendlyName();
                _safeSerialPort.Open();
                openSucceeded = _safeSerialPort.IsOpen;
            }
            catch (IOException e)
            {
                if (showErrorPopup)
                {
                    Common.ShowErrorMessageBox(e, $"Failed to open port {SerialPortSetting.ComPort}.");
                }
                Logger.Error(e);
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Error);
            }
            catch (Exception e)
            {
                if (showErrorPopup)
                {
                    Common.ShowErrorMessageBox(e, $"Failed to open port {SerialPortSetting.ComPort}.");
                }
                Logger.Error(e);
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Error);
            }

            _portShouldBeOpen = openSucceeded;

            if (openSucceeded)
            {
                Logger.Info($"Serial port opened {SerialPortSetting.ComPort}. BaudRate={SerialPortSetting.BaudRate}, WriteTimeout={SerialPortSetting.WriteTimeout}, ReadTimeout={SerialPortSetting.ReadTimeout}");
                _ = Task.Run(SerialDataWrite);
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Opened);
            }
        }

        public bool IsOpen => _safeSerialPort != null && _safeSerialPort.IsOpen;

        public void Close()
        {
            try
            {
                if (_safeSerialPort == null) return;

                _portShouldBeOpen = false;
                _serialDataWaitingForWriteResetEvent?.Set();

                Logger.Info($"Closing and disposing serial port {SerialPortSetting.ComPort}");

                _serialReceiver?.Dispose();
                _serialReceiver = null;

                _safeSerialPort.Close();
                _safeSerialPort.Dispose();
                _safeSerialPort = null;

                ClearQueuedData("port-close");

                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Closed);
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }
        }

        public void DcsBiosBulkDataReceived(object sender, DCSBIOSBulkDataEventArgs e)
        {
            try
            {
                QueueSerialData(e.Data);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        public void ApplyPortConfig(SerialPortSetting serialPortSetting)
        {
            if(serialPortSetting == null) return;

            var wasOpen = _safeSerialPort != null && _safeSerialPort.IsOpen;

            SerialPortSetting = serialPortSetting;
            Close();

            if (wasOpen)
            {
                Open();
            }
        }

        private void ApplyPortConfig()
        {
            if (_safeSerialPort == null) return;
            
            _safeSerialPort.PortName = SerialPortSetting.ComPort;
            _safeSerialPort.BaudRate = SerialPortSetting.BaudRate;
            _safeSerialPort.Parity = SerialPortSetting.Parity;
            _safeSerialPort.StopBits = SerialPortSetting.Stopbits;
            _safeSerialPort.DataBits = SerialPortSetting.Databits;
            _safeSerialPort.Handshake = SerialPortSetting.Handshake;
            _safeSerialPort.DtrEnable = SerialPortSetting.LineSignalDtr;
            _safeSerialPort.RtsEnable = SerialPortSetting.LineSignalRts;
            _safeSerialPort.WriteTimeout = SerialPortSetting.WriteTimeout == 0 ? SerialPort.InfiniteTimeout : SerialPortSetting.WriteTimeout;
            _safeSerialPort.ReadTimeout = SerialPortSetting.ReadTimeout == 0 ? SerialPort.InfiniteTimeout : SerialPortSetting.ReadTimeout;

            GetFriendlyName();
        }

        private void QueueSerialData(byte[] data)
        {
            if (data == null || data.Length == 0 || _safeSerialPort == null || !_safeSerialPort.IsOpen) return;

            _serialDataQueue.Enqueue(data);
            Interlocked.Increment(ref _queuedMessageCount);
            Interlocked.Add(ref _queuedByteCount, data.Length);
            LogQueueDiagnosticsIfNeeded("enqueue");
            _serialDataWaitingForWriteResetEvent.Set();
        }

        private void SerialDataWrite()
        {
            while (true)
            {
                try
                {
                    _serialDataWaitingForWriteResetEvent.WaitOne();
                    if (_shutdown || _safeSerialPort == null || !_safeSerialPort.IsOpen) break;

                    while (_serialDataQueue.TryDequeue(out var data))
                    {
                        if (_shutdown || _safeSerialPort == null || !_safeSerialPort.IsOpen)
                        {
                            break;
                        }

                        Interlocked.Decrement(ref _queuedMessageCount);
                        Interlocked.Add(ref _queuedByteCount, -data.Length);
                        NormalizeQueueCounters();

                        _safeSerialPort.Write(data, 0, data.Length);
                        _lastSerialWriteActivityUtc = DateTime.UtcNow;
                        DBEventManager.BroadCastSerialData(ComPort, data.Length, StreamInterface.SerialPortWritten);
                    }

                    LogQueueDiagnosticsIfNeeded("dequeue");
                }
                catch (TimeoutException e)
                {
                    Logger.Error("SerialDataWrite failed => {0}", e);
                    _portShouldBeOpen = false;
                    ClearQueuedData("serial-write-timeout");
                    DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.TimeOutError);
                    break;
                }
                catch (OperationCanceledException e)
                {
                    Logger.Error("SerialDataWrite failed => {0}", e);
                    _portShouldBeOpen = false;
                    ClearQueuedData("serial-write-cancelled");
                    DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Error);
                    break;
                }
                catch (IOException e)
                {
                    Logger.Error("SerialDataWrite failed => {0}", e);
                    _portShouldBeOpen = false;
                    ClearQueuedData("serial-write-ioerror");
                    DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.IOError);
                    break;
                }
                catch (Exception e)
                {
                    Logger.Error("SerialDataWrite failed => {0}", e);
                    _portShouldBeOpen = false;
                    ClearQueuedData("serial-write-exception");
                    DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Error);
                    break;
                }
            }
        }

        private void LogQueueDiagnosticsIfNeeded(string reason, bool force = false)
        {
            var queuedMessages = Volatile.Read(ref _queuedMessageCount);
            var queuedBytes = Volatile.Read(ref _queuedByteCount);
            var queueWarning = queuedMessages >= QueueMessageWarningThreshold || queuedBytes >= QueueBytesWarningThreshold;

            if (!force && !queueWarning)
            {
                return;
            }

            var utcNow = DateTime.UtcNow;
            if (!force && utcNow - _lastQueueDiagnosticsLogUtc < QueueDiagnosticsLogInterval)
            {
                return;
            }

            _lastQueueDiagnosticsLogUtc = utcNow;

            if (queueWarning)
            {
                Logger.Warn($"Serial queue pressure on {SerialPortSetting.ComPort}. Reason={reason}, QueuedMessages={queuedMessages}, QueuedBytes={queuedBytes}");
                return;
            }

            Logger.Info($"Serial queue diagnostic on {SerialPortSetting.ComPort}. Reason={reason}, QueuedMessages={queuedMessages}, QueuedBytes={queuedBytes}");
        }

        private int ClearQueuedData(string reason)
        {
            var clearedMessages = 0;
            long clearedBytes = 0;

            while (_serialDataQueue.TryDequeue(out var queuedData))
            {
                clearedMessages++;
                clearedBytes += queuedData?.Length ?? 0;
            }

            Interlocked.Add(ref _queuedMessageCount, -clearedMessages);
            Interlocked.Add(ref _queuedByteCount, -clearedBytes);
            NormalizeQueueCounters();

            if (clearedMessages > 0 || clearedBytes > 0)
            {
                Logger.Info($"Cleared queued serial data on {SerialPortSetting.ComPort}. Reason={reason}, ClearedMessages={clearedMessages}, ClearedBytes={clearedBytes}");
            }

            return clearedMessages;
        }

        private void NormalizeQueueCounters()
        {
            if (Volatile.Read(ref _queuedMessageCount) < 0)
            {
                Interlocked.Exchange(ref _queuedMessageCount, 0);
            }

            if (Volatile.Read(ref _queuedByteCount) < 0)
            {
                Interlocked.Exchange(ref _queuedByteCount, 0);
            }
        }

        private static IEnumerable<RegistryKey> GetSubKeys(RegistryKey key)
        {
            foreach (var keyName in key.GetSubKeyNames())
                using (var subKey = key.OpenSubKey(keyName))
                    yield return subKey;
        }

        private static string GetName(RegistryKey key)
        {
            var name = key.Name;
            int idx;
            return (idx = name.LastIndexOf('\\')) == -1 ? name : name[(idx + 1)..];
        }

        private void GetFriendlyName()
        {
            if (!GetFriendlyName1())
            {
                GetFriendlyName2();
            }
        }

        private bool GetFriendlyName1()
        {
            using var usbDevicesKey = Registry.LocalMachine.OpenSubKey(Constants.USBDevices);

            foreach (var usbDeviceKey in GetSubKeys(usbDevicesKey))
            {
                foreach (var devFnKey in GetSubKeys(usbDeviceKey))
                {
                    var friendlyName = (string)devFnKey.GetValue("FriendlyName") ?? (string)devFnKey.GetValue("DeviceDesc");

                    using var deviceParametersKey = devFnKey.OpenSubKey("Device Parameters");
                    var portName = (string)deviceParametersKey?.GetValue("PortName");

                    if (string.IsNullOrEmpty(portName) || SerialPortSetting.ComPort != portName) continue;

                    FriendlyName = friendlyName?.Replace($"({SerialPortSetting.ComPort})", "", StringComparison.Ordinal);
                    VIDPID = GetName(usbDeviceKey);
                    FriendlyName = string.IsNullOrEmpty(FriendlyName) ? VIDPID : FriendlyName;

                    return true;
                    //yield return new UsbSerialPort(portName, GetName(devBaseKey) + @"\" + GetName(devFnKey), friendlyName);
                }
            }

            return false;
        }

        private bool GetFriendlyName2()
        {
            using var devicesKeys = Registry.LocalMachine.OpenSubKey(Constants.DeviceEnumeration);

            foreach (var deviceKey in GetSubKeys(devicesKeys))
            {
                foreach (var deviceSub1Key in GetSubKeys(deviceKey))
                {
                    foreach (var deviceSub2Key in GetSubKeys(deviceSub1Key))
                    {
                        var friendlyName = (string)deviceSub2Key.GetValue("FriendlyName") ?? (string)deviceSub2Key.GetValue("DeviceDesc");

                        using var deviceParametersKey = deviceSub2Key.OpenSubKey("Device Parameters");
                        var portName = (string)deviceParametersKey?.GetValue("PortName");

                        if (string.IsNullOrEmpty(portName) || SerialPortSetting.ComPort != portName) continue;

                        FriendlyName = friendlyName?.Replace($"({SerialPortSetting.ComPort})", "", StringComparison.Ordinal);
                        VIDPID = GetName(deviceKey);
                        FriendlyName = string.IsNullOrEmpty(FriendlyName) ? VIDPID : FriendlyName;

                        return true;
                    }
                }
            }

            return false;
        }

        public Handshake Handshake
        {
            get => SerialPortSetting.Handshake;
            set
            {
                SerialPortSetting.Handshake = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public string ComPort
        {
            get => SerialPortSetting.ComPort;
            set
            {
                SerialPortSetting.ComPort = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public string FriendlyName { get; set; }

        public string VIDPID { get; set; }

        public int BaudRate
        {
            get => SerialPortSetting.BaudRate;
            set
            {
                SerialPortSetting.BaudRate = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public int Databits
        {
            get => SerialPortSetting.Databits;
            set
            {
                SerialPortSetting.Databits = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public StopBits Stopbits
        {
            get => SerialPortSetting.Stopbits;
            set
            {
                SerialPortSetting.Stopbits = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public Parity Parity
        {
            get => SerialPortSetting.Parity;
            set
            {
                SerialPortSetting.Parity = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public int WriteTimeout
        {
            get => SerialPortSetting.WriteTimeout;
            set
            {
                SerialPortSetting.WriteTimeout = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public int ReadTimeout
        {
            get => SerialPortSetting.ReadTimeout;
            set
            {
                SerialPortSetting.ReadTimeout = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public bool LineSignalDtr
        {
            get => SerialPortSetting.LineSignalDtr;
            set
            {
                SerialPortSetting.LineSignalDtr = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public bool LineSignalRts
        {
            get => SerialPortSetting.LineSignalRts;
            set
            {
                SerialPortSetting.LineSignalRts = value;
                DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Settings, 0, null, SerialPortSetting);
            }
        }

        public static bool SerialPortCurrentlyExists(string portName)
        {
            if (string.IsNullOrEmpty(portName))
            {
                return false;
            }
            var existingPorts = Common.GetSerialPortNames();
            return existingPorts.Any(portName.Equals);
        }

        public void OnDataReceived(SerialDataEventArgs e)
        {
            if (e.ComPort != ComPort)
            {
                return;
            }

            if (e.StreamInterface == StreamInterface.SerialPortRead)
            {
                _lastSerialReadActivityUtc = DateTime.UtcNow;
                _readChannelObserved = true;
                return;
            }

            if (e.StreamInterface == StreamInterface.SerialPortWritten)
            {
                _lastSerialWriteActivityUtc = DateTime.UtcNow;
            }
        }

        private bool ShouldWatchDogBark(DateTime utcNow)
        {
            if (!Settings.Default.WatchDogEnabled)
            {
                return false;
            }

            var watchDogNoReadTimeoutSeconds = Settings.Default.WatchDogNoReadTimeoutSeconds;
            var watchDogRecentWriteWindowSeconds = Settings.Default.WatchDogRecentWriteWindowSeconds;
            var watchDogNoReadTimeout = TimeSpan.FromSeconds(Math.Max(1, watchDogNoReadTimeoutSeconds));
            var watchDogRecentWriteWindow = TimeSpan.FromSeconds(Math.Max(1, watchDogRecentWriteWindowSeconds));
            var watchDogCooldown = TimeSpan.FromSeconds(Math.Max(1, Settings.Default.WatchDogCooldownSeconds));

            if (_safeSerialPort == null || !_safeSerialPort.IsOpen)
            {
                return false;
            }

            if (!_readChannelObserved)
            {
                return false;
            }

            if (_lastSerialWriteActivityUtc == DateTime.MinValue || _lastSerialReadActivityUtc == DateTime.MinValue)
            {
                return false;
            }

            if (utcNow - _lastWatchDogBarkUtc < watchDogCooldown)
            {
                return false;
            }

            if (watchDogRecentWriteWindowSeconds > 0 && utcNow - _lastSerialWriteActivityUtc > watchDogRecentWriteWindow)
            {
                return false;
            }

            if (watchDogNoReadTimeoutSeconds <= 0)
            {
                return false;
            }

            return utcNow - _lastSerialReadActivityUtc > watchDogNoReadTimeout;
        }

        private void BarkWatchDog()
        {
            var utcNow = DateTime.UtcNow;
            var watchDogNoReadTimeoutSeconds = Settings.Default.WatchDogNoReadTimeoutSeconds;
            var watchDogRecentWriteWindowSeconds = Settings.Default.WatchDogRecentWriteWindowSeconds;
            var watchDogCooldownSeconds = Math.Max(1, Settings.Default.WatchDogCooldownSeconds);
            var watchDogReopenDelayMs = Math.Max(1, Settings.Default.WatchDogReopenDelayMilliseconds);

            var secondsSinceLastRead = _lastSerialReadActivityUtc == DateTime.MinValue
                ? -1
                : (utcNow - _lastSerialReadActivityUtc).TotalSeconds;
            var secondsSinceLastWrite = _lastSerialWriteActivityUtc == DateTime.MinValue
                ? -1
                : (utcNow - _lastSerialWriteActivityUtc).TotalSeconds;

            _lastWatchDogBarkUtc = utcNow;
            _watchDogBarkCount++;

            Logger.Warn($"Watchdog bark on {SerialPortSetting.ComPort}. Reopening serial port. BarkCount={_watchDogBarkCount}, SecondsSinceLastRead={secondsSinceLastRead:F1}, SecondsSinceLastWrite={secondsSinceLastWrite:F1}, NoReadTimeoutSeconds={watchDogNoReadTimeoutSeconds}, RecentWriteWindowSeconds={watchDogRecentWriteWindowSeconds}, CooldownSeconds={watchDogCooldownSeconds}, ReopenDelayMilliseconds={watchDogReopenDelayMs}, ReadTimerDisabled={watchDogNoReadTimeoutSeconds <= 0}, WriteTimerDisabled={watchDogRecentWriteWindowSeconds <= 0}");
            DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.WatchDogBark);

            Close();
            Thread.Sleep(watchDogReopenDelayMs);
            Open(false);
        }

        private void CheckPortOpen()
        {
            try
            {
                while (!_shutdown)
                {
                    if (_portShouldBeOpen)
                    {
                        if (_safeSerialPort == null || !_safeSerialPort.IsOpen)
                        {
                            Logger.Error("Background Thread (CheckPortOpen) detected port is not open.");
                            DBEventManager.BroadCastPortStatus(SerialPortSetting.ComPort, SerialPortStatus.Critical);
                            break;
                        }

                        if (ShouldWatchDogBark(DateTime.UtcNow))
                        {
                            BarkWatchDog();
                        }
                    }

                    Thread.Sleep(Constants.MS1000);
                }
            }
            catch (Exception e)
            {
                Logger.Error(e, "Port check thread.");
            }
        }
    }
}
