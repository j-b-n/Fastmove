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
            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);           
            ListBox1.ItemsSource = Globals.ThisAddIn._items;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ThemeManager themeManager = new ThemeManager(this);
            themeManager.ChangeThemeIfWindowsChangedIt();
            //WatchTheme();
            //ChangeThemeIfWindowsChangedIt();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {            
            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void AboutBtn_Click(object sender, RoutedEventArgs e)
        {                      
            WPF.AboutWindow ui = new WPF.AboutWindow();
            ui.Show();
        }
    }    
}
