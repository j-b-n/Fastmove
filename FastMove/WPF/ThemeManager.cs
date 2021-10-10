using AdonisUI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FastMove.WPF
{
    public class ThemeManager
    {
        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private const string RegistryValueName = "AppsUseLightTheme";

        private Uri _currentTheme = null;

        AdonisUI.Controls.AdonisWindow parent = null;

        public ThemeManager(AdonisUI.Controls.AdonisWindow ParentWindow)
        {
            parent = ParentWindow;
        }

        private void ChangeTheme(Uri theme)
        {
            if (theme == ResourceLocator.DarkColorScheme)
            {
                ResourceLocator.SetColorScheme(parent.Resources, ResourceLocator.DarkColorScheme, ResourceLocator.LightColorScheme);
            }
            else
            {
                ResourceLocator.SetColorScheme(parent.Resources, ResourceLocator.LightColorScheme, ResourceLocator.DarkColorScheme);
            }
        }

        private static Uri GetWindowsTheme()
        {
            var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            var registryValueObject = key?.GetValue(RegistryValueName);
            if (registryValueObject == null)
            {
                return ResourceLocator.LightColorScheme;
            }

            var registryValue = (int)registryValueObject;

            return registryValue > 0 ? ResourceLocator.LightColorScheme : ResourceLocator.DarkColorScheme;
        }

        public void ChangeThemeIfWindowsChangedIt()
        {
            var newWindowsTheme = GetWindowsTheme();
            if (_currentTheme != newWindowsTheme)
            {
                _currentTheme = newWindowsTheme;
                ChangeTheme(_currentTheme);
            }
        }
    }
}