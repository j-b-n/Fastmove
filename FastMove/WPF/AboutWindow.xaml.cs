using AdonisUI;
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
    /// Interaction logic for AboutWindow.xaml
    /// </summary>
    public partial class AboutWindow : AdonisUI.Controls.AdonisWindow
    {
        public AboutWindow()
        {
            InitializeComponent();
            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);
            ThemeManager themeManager = new ThemeManager(this);
            themeManager.ChangeThemeIfWindowsChangedIt();            
        }
    }
}
