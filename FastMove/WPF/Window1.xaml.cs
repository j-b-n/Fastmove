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
        ThemeManager themeManager = null;

        public Window1()
        {
            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);
            InitializeComponent();                                
        }

        public string Pad(int i)
        {
            if (i < 10)
            {
                return "0" + i;
            }
            return i + "";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            themeManager = new ThemeManager(this);
            themeManager.SetDefaultTheme();

            ListBox1.ItemsSource = Globals.ThisAddIn._items;
            RecentListBox.ItemsSource = Globals.ThisAddIn._recentItems;

            TextMoveMail.Text = "Move mail to (" + ListBox1.Items.Count + "):";
            if (ListBox1.Items.Count > 0)
                ListBox1.SelectedItem = ListBox1.Items.GetItemAt(0);
            
            double seconds = Globals.ThisAddIn._InboxAvg;
            TimeSpan TS = TimeSpan.FromSeconds(seconds);
            string AvgText = TS.Days + " days," +
                Pad(TS.Hours) + " hours, " +
                Pad(TS.Minutes) + " minutes " +
                Pad(TS.Seconds) + " seconds";
            AvgTimeInboxTextBox.Text = AvgText;

            seconds = 0;
            int count = 0;
            foreach (double d in Globals.ThisAddIn._avgTimeBeforeMove)
            {
                seconds += d;
                count++;
            }
            if (count > 0)
                seconds /= count;

            TS = TimeSpan.FromSeconds(seconds);
            AvgText = TS.Days + " days," +
                Pad(TS.Hours) + " hours, " +
                Pad(TS.Minutes) + " minutes " +
                Pad(TS.Seconds) + " seconds";
            AvgTimeTextBox.Text = AvgText;

            DeferCheckBox.IsChecked = Globals.ThisAddIn._deferEmails;
            if (Globals.ThisAddIn._deferEmails)
            {
                DateTime deferTime = Globals.ThisAddIn.NextPossibleSendTime();
                DeferTextBox.Text = string.Format("{0}", deferTime.ToString());
            }
            else
            {
                DeferTextBox.Text = string.Format("Immediately");
            }

            ///Update StatusBar content

            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            DateTime day = DateTime.Today.Date;
            count = 0;

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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            object selectedItem = ListBox1.SelectedItem;
            
            MessageBox.Show(selectedItem.ToString());
            //Globals.ThisAddIn.MoveMail(selectedItem.ToString());
            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }        

        private bool Compare(string s)
        {
            string t = TextBox1.Text;

            s = s.ToLower();
            t = t.ToLower();

            if (s.Contains(t))
            {
                return true;
            }
            return false;
        }

        private void TextBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBox1.Text.Length > 0)
            {
                List<string> Searchitems = Globals.ThisAddIn._items.FindAll(Compare);
                List<string> _Searchitems = new List<string>();

                Searchitems.Sort();

                string t = TextBox1.Text;
                t = t.ToLower();
                bool addFirst = false;

                foreach (string item in Searchitems)
                {
                    string _item = item.ToLower();
                    //If we have an exact match for a part in the path of a folder - return true
                    string[] words = _item.Split('\\');
                    addFirst = false;
                    foreach (string word in words)
                    {
                        if (word.Equals(t))
                        {
                            addFirst = true;
                            break;
                        }
                    }

                    if (addFirst)
                        _Searchitems.Insert(0, item);
                    else
                        _Searchitems.Add(item);
                }

                ListBox1.ItemsSource = _Searchitems;
                if(ListBox1.Items.Count > 0)
                    ListBox1.SelectedItem = ListBox1.Items.GetItemAt(0);

            }
            else
                ListBox1.ItemsSource = Globals.ThisAddIn._items;
            
            TextMoveMail.Text = "Move mail to (" + ListBox1.Items.Count + "):";
        }

        private void DeferCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn._deferEmails = (bool)DeferCheckBox.IsChecked;

            if (Globals.ThisAddIn._deferEmails)
            {
                DateTime deferTime = Globals.ThisAddIn.NextPossibleSendTime();
                DeferTextBox.Text = string.Format("{0}", deferTime.ToString());
            }
            else
            {
                DeferTextBox.Text = string.Format("Immediately");
            }
        }

        private void ToogleThemeBtn_Click(object sender, RoutedEventArgs e)
        {                     
            themeManager.ToggleTheme();
            if (themeManager.GetTheme() == ResourceLocator.DarkColorScheme)
                ToogleThemeBtn.Content = "☀️";
            else
                ToogleThemeBtn.Content = "🌙";
        }

        //Buttons
        private void DeferBtn_Click(object sender, RoutedEventArgs e)
        {
            WPF.DeferWindow ui = new WPF.DeferWindow();
            ui.Show();
        }
        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            WPF.RefreshWindow ui = new WPF.RefreshWindow();
            ui.Show();
        }        
        private void StatisticsBtn_Click(object sender, RoutedEventArgs e)
        {
            WPF.StatisticsWindow ui = new WPF.StatisticsWindow();
            ui.Show();
        }
        private void SettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            WPF.SettingsWindow ui = new WPF.SettingsWindow();
            ui.Show();
        }
        private void AboutBtn_Click(object sender, RoutedEventArgs e)
        {
            WPF.AboutWindow ui = new WPF.AboutWindow();
            ui.Show();
        }        
    }    
}
