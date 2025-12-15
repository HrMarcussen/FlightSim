using Microsoft.Win32;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FlightSimTool
{
    public static class ThemeManager
    {
        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        private const string RegistryValueName = "AppsUseLightTheme";

        public enum Theme
        {
            Light,
            Dark
        }

        public static void Initialize()
        {
            // Listen to system settings changes
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;

            // Initial apply
            ApplyTheme(GetCurrentSystemTheme());
        }

        private static void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
            {
                Application.Current.Dispatcher.Invoke(() => ApplyTheme(GetCurrentSystemTheme()));
            }
        }

        public static Theme GetCurrentSystemTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    object registryValueObject = key?.GetValue(RegistryValueName);
                    if (registryValueObject == null)
                    {
                        return Theme.Light; // Default fallback
                    }

                    int registryValue = (int)registryValueObject;
                    return registryValue > 0 ? Theme.Light : Theme.Dark;
                }
            }
            catch
            {
                return Theme.Light;
            }
        }

        public static void ApplyTheme(Theme theme)
        {
            string themeUri = theme == Theme.Light 
                ? "Themes/LightTheme.xaml" 
                : "Themes/DarkTheme.xaml";

            try 
            {
                var dict = new ResourceDictionary { Source = new Uri(themeUri, UriKind.Relative) };

                // Clear existing theme dictionaries (assuming theme dict is the first one or we manage it specifically)
                // For simplicity, we'll clear and re-add or find by source if we had multiple
                
                // Strategy: Find existing theme dictionary if present and replace it, otherwise add it.
                // Since this is a simple app, we can probably just assume MergedDictionaries[0] is our theme 
                // IF we set it up that way in App.xaml.
                
                var app = Application.Current;
                if (app == null) return;

                // Better safe strategy: Remove any dict that looks like a theme, then add the new one
                for (int i = app.Resources.MergedDictionaries.Count - 1; i >= 0; i--)
                {
                    var md = app.Resources.MergedDictionaries[i];
                    if (md.Source != null && md.Source.ToString().Contains("Themes/") && !md.Source.ToString().Contains("Styles.xaml"))
                    {
                        app.Resources.MergedDictionaries.RemoveAt(i);
                    }
                }

                app.Resources.MergedDictionaries.Add(dict);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to apply theme: {ex.Message}");
            }
        }
    }
}
