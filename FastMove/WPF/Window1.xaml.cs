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
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : AdonisUI.Controls.AdonisWindow
    {
        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private const string RegistryValueName = "AppsUseLightTheme";

        private Uri _currentTheme = null;

        public Window1()
        {            
            InitializeComponent();
            ChangeThemeIfWindowsChangedIt();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //WatchTheme();
            //ChangeThemeIfWindowsChangedIt();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {            
            this.Close();
        }

        private  void ChangeTheme(Uri theme)
        {            
            if (theme == ResourceLocator.DarkColorScheme)
            {                
                ResourceLocator.SetColorScheme(this.Resources, ResourceLocator.DarkColorScheme, ResourceLocator.LightColorScheme);         
            }
            else
            {
                ResourceLocator.SetColorScheme(this.Resources, ResourceLocator.LightColorScheme, ResourceLocator.DarkColorScheme);                
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

        private void ChangeThemeIfWindowsChangedIt()
        {
            var newWindowsTheme = GetWindowsTheme();            
            if (_currentTheme != newWindowsTheme)
            {
                _currentTheme = newWindowsTheme;
                ChangeTheme(_currentTheme);
            }
        }
    }

    static class ResourceLocator
    {
        public static Uri DarkColorScheme => new Uri("pack://application:,,,/AdonisUI;component/ColorSchemes/Dark.xaml", UriKind.Absolute);

        public static Uri LightColorScheme => new Uri("pack://application:,,,/AdonisUI;component/ColorSchemes/Light.xaml", UriKind.Absolute);

        public static Uri ClassicTheme => new Uri("pack://application:,,,/AdonisUI.ClassicTheme;component/Resources.xaml", UriKind.Absolute);

        /// <summary> Adds any Adonis theme to the provided resource dictionary. </summary>
        /// <param name="rootResourceDictionary">
        /// The resource dictionary containing AdonisUI's resources. Expected are the resource
        /// dictionaries of the app or window.
        /// </param>
        public static void AddAdonisResources(ResourceDictionary rootResourceDictionary)
        {
            rootResourceDictionary.MergedDictionaries.Add(new ResourceDictionary { Source = ClassicTheme });
        }

        /// <summary> Removes all resources of AdonisUI from the provided resource dictionary. </summary>
        /// <param name="rootResourceDictionary">
        /// The resource dictionary containing AdonisUI's resources. Expected are the resource
        /// dictionaries of the app or window.
        /// </param>
        public static void RemoveAdonisResources(ResourceDictionary rootResourceDictionary)
        {
            Uri[] adonisResources = { ClassicTheme };
            var currentTheme = FindFirstContainedResourceDictionaryByUri(rootResourceDictionary, adonisResources);

            if (currentTheme != null)
            {
                RemoveResourceDictionaryFromResourcesDeep(currentTheme, rootResourceDictionary);
            }
        }

        /// <summary>
        /// Adds a resource dictionary with the specified uri to the MergedDictionaries
        /// collection of the <see cref="rootResourceDictionary" />. Additionally all child
        /// ResourceDictionaries are traversed recursively to find the current color scheme
        /// which is removed if found.
        /// </summary>
        /// <param name="rootResourceDictionary">
        /// The resource dictionary containing the currently active color scheme. It will
        /// receive the new color scheme in its MergedDictionaries. Expected are the resource
        /// dictionaries of the app or window.
        /// </param>
        /// <param name="colorSchemeResourceUri">
        /// The Uri of the color scheme to be set. Can be taken from the
        /// <see cref="ResourceLocator" /> class.
        /// </param>
        /// <param name="currentColorSchemeResourceUri">
        /// Optional uri to an external color scheme that is not provided by AdonisUI.
        /// </param>
        public static void SetColorScheme(ResourceDictionary rootResourceDictionary, Uri colorSchemeResourceUri, Uri currentColorSchemeResourceUri)
        {
            var knownColorSchemes = currentColorSchemeResourceUri != null ? new[] { currentColorSchemeResourceUri } : new[] { LightColorScheme, DarkColorScheme };

            var currentTheme = FindFirstContainedResourceDictionaryByUri(rootResourceDictionary, knownColorSchemes);

            if (currentTheme != null)
            {
                RemoveResourceDictionaryFromResourcesDeep(currentTheme, rootResourceDictionary);
            }

            rootResourceDictionary.MergedDictionaries.Add(new ResourceDictionary { Source = colorSchemeResourceUri });
        }

        private static ResourceDictionary FindFirstContainedResourceDictionaryByUri(ResourceDictionary resourceDictionary, Uri[] knownColorSchemes)
        {
            if (knownColorSchemes.Any(scheme => resourceDictionary.Source != null && resourceDictionary.Source.IsAbsoluteUri && resourceDictionary.Source.AbsoluteUri.Equals(scheme.AbsoluteUri, StringComparison.InvariantCulture)))
            {
                return resourceDictionary;
            }

            if (!resourceDictionary.MergedDictionaries.Any())
            {
                return null;
            }

            return resourceDictionary.MergedDictionaries.FirstOrDefault(d => FindFirstContainedResourceDictionaryByUri(d, knownColorSchemes) != null);
        }

        private static bool RemoveResourceDictionaryFromResourcesDeep(ResourceDictionary resourceDictionaryToRemove, ResourceDictionary rootResourceDictionary)
        {
            if (!rootResourceDictionary.MergedDictionaries.Any())
            {
                return false;
            }

            if (rootResourceDictionary.MergedDictionaries.Contains(resourceDictionaryToRemove))
            {
                rootResourceDictionary.MergedDictionaries.Remove(resourceDictionaryToRemove);
                return true;
            }

            return rootResourceDictionary.MergedDictionaries.Any(dict => RemoveResourceDictionaryFromResourcesDeep(resourceDictionaryToRemove, dict));
        }
    }
}
