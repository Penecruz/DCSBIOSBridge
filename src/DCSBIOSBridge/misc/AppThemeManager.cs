using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using DCSBIOSBridge.Properties;

namespace DCSBIOSBridge.misc
{
    public enum ThemeMode
    {
        FollowWindows = 0,
        Light = 1,
        Dark = 2
    }

    public static class AppThemeManager
    {
        private const string LightThemePath = "Themes/LightTheme.xaml";
        private const string DarkThemePath = "Themes/DarkTheme.xaml";
        private const string LightThemePathNormalized = "/themes/lighttheme.xaml";
        private const string DarkThemePathNormalized = "/themes/darktheme.xaml";
        private const string PersonalizeRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        private const string AppsUseLightThemeValueName = "AppsUseLightTheme";
        private const int DwmUseImmersiveDarkModeAttribute = 20;
        private static bool _initialized;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        public static void Initialize()
        {
            ApplyConfiguredTheme();

            if (_initialized)
            {
                return;
            }

            SystemEvents.UserPreferenceChanged += SystemEventsOnUserPreferenceChanged;
            _initialized = true;
        }

        public static void Shutdown()
        {
            if (!_initialized)
            {
                return;
            }

            SystemEvents.UserPreferenceChanged -= SystemEventsOnUserPreferenceChanged;
            _initialized = false;
        }

        public static ThemeMode GetConfiguredThemeMode()
        {
            return Enum.IsDefined(typeof(ThemeMode), Settings.Default.ThemeMode)
                ? (ThemeMode)Settings.Default.ThemeMode
                : ThemeMode.FollowWindows;
        }

        public static void ApplyConfiguredTheme()
        {
            ApplyTheme(GetConfiguredThemeMode());
        }

        public static void ApplyTheme(ThemeMode themeMode)
        {
            var useDark = ShouldUseDarkTheme(themeMode);

            var targetThemePath = useDark ? DarkThemePath : LightThemePath;
            var mergedDictionaries = Application.Current.Resources.MergedDictionaries;

            for (var i = mergedDictionaries.Count - 1; i >= 0; i--)
            {
                var source = mergedDictionaries[i].Source?.ToString();
                if (string.IsNullOrEmpty(source))
                {
                    continue;
                }

                var normalizedSource = source.Replace('\\', '/').ToLowerInvariant();
                if (normalizedSource.EndsWith(LightThemePathNormalized, StringComparison.Ordinal) ||
                    normalizedSource.EndsWith(DarkThemePathNormalized, StringComparison.Ordinal))
                {
                    mergedDictionaries.RemoveAt(i);
                }
            }

            mergedDictionaries.Add(new ResourceDictionary
            {
                Source = new Uri(targetThemePath, UriKind.Relative)
            });

            foreach (Window window in Application.Current.Windows)
            {
                ApplyTitleBarTheme(window, useDark);
            }
        }

        public static void ApplyTitleBarTheme(Window window)
        {
            if (window == null)
            {
                return;
            }

            ApplyTitleBarTheme(window, ShouldUseDarkTheme(GetConfiguredThemeMode()));
        }

        private static bool IsWindowsDarkMode()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(PersonalizeRegistryKey);
                var value = key?.GetValue(AppsUseLightThemeValueName);

                if (value is int intValue)
                {
                    return intValue == 0;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool ShouldUseDarkTheme(ThemeMode themeMode)
        {
            return themeMode switch
            {
                ThemeMode.Dark => true,
                ThemeMode.Light => false,
                _ => IsWindowsDarkMode()
            };
        }

        private static void ApplyTitleBarTheme(Window window, bool useDark)
        {
            try
            {
                var windowHandle = new WindowInteropHelper(window).Handle;
                if (windowHandle == IntPtr.Zero)
                {
                    return;
                }

                var enabled = useDark ? 1 : 0;
                _ = DwmSetWindowAttribute(windowHandle, DwmUseImmersiveDarkModeAttribute, ref enabled, sizeof(int));
            }
            catch
            {
            }
        }

        private static void SystemEventsOnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (GetConfiguredThemeMode() != ThemeMode.FollowWindows)
            {
                return;
            }

            if (e.Category != UserPreferenceCategory.General &&
                e.Category != UserPreferenceCategory.Color &&
                e.Category != UserPreferenceCategory.VisualStyle)
            {
                return;
            }

            Application.Current?.Dispatcher.BeginInvoke(() => ApplyTheme(ThemeMode.FollowWindows));
        }
    }
}
