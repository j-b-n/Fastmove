﻿using AdonisUI;
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
        ThemeManager themeManager = null;

        public AboutWindow()
        {
            InitializeComponent();
            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);
            themeManager = new ThemeManager(this);
            themeManager.SetDefaultTheme();
        }
        private void ToogleThemeBtn_Click(object sender, RoutedEventArgs e)
        {
            themeManager.ToggleTheme();

            if (themeManager.GetTheme() == ResourceLocator.DarkColorScheme)
                ToogleThemeBtn.Content = "☀️";
            else
                ToogleThemeBtn.Content = "🌙";

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateInfo ui = new UpdateInfo();

            int AddinUpdateAvailable = ui.CheckForUpdate();
            string runningVersion = Globals.ThisAddIn.publishedVersion;

            ThisVersionTB.Text = string.Format("Version: '{0}'", runningVersion);
            OnlineVersionTB.Text = string.Format("Online version: '{0}'", ui.UpdateVariables.Version);

            ///Update StatusBar content

            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            DateTime day = DateTime.Today.Date;
            int count = 0;

            if (_MailsPerDay.ContainsKey(day))
            {
                count = _MailsPerDay[day];
            }
            SBToday.Text = string.Format("Today: {0}", count);

            //Last week
            count = 0;

            for (int i = 0; i > -6; i--)
            {
                if (_MailsPerDay.ContainsKey(day))
                {
                    count += _MailsPerDay[day];
                }
                day = day.AddDays(i);
            }

            SBLastWeek.Text = string.Format("Last week: {0}", count);

            SBVersion.Text = string.Format("Version: {0}", Globals.ThisAddIn.publishedVersion);

        }
    }
}
