using System;
using System.Collections.Generic;
using System.ComponentModel;
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

    public class ComboData
    {
        public int Id { get; set; }
        public string Value { get; set; }
    }

    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : AdonisUI.Controls.AdonisWindow
    {                
        private readonly List<ComboData> ListData = new List<ComboData> {
            new ComboData { Id = 30, Value = "30 minutes" },
            new ComboData { Id = 60, Value = "1 hour" },
            new ComboData { Id = 120, Value = "2 hours" },
            new ComboData { Id = 1440, Value = "24 hours" },
            new ComboData { Id = 10080, Value = "1 week" },
            new ComboData { Id = 20160, Value = "2 weeks" }
        };

        private readonly BindingList<string> _ignoreL = new BindingList<string>();
        private readonly BindingList<string> _FolderLevel1L = new BindingList<string>();
        List<string> _originalIgnoreList = new List<string>();

        ThemeManager themeManager = null;

        public SettingsWindow()
        {
            InitializeComponent();
            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);
            themeManager = new ThemeManager(this);
            themeManager.SetDefaultTheme();        
            
            UpdateCB.ItemsSource = ListData;
            UpdateCB.DisplayMemberPath = "Value";
            UpdateCB.SelectedValuePath = "Id";            

        }
        private void ToogleThemeBtn_Click(object sender, RoutedEventArgs e)
        {
            themeManager.ToggleTheme();

            if (themeManager.GetTheme() == ResourceLocator.DarkColorScheme)
                ToogleThemeBtn.Content = "☀️";
            else
                ToogleThemeBtn.Content = "🌙";

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {            
            List<string> _FolderLevel1List;

            _originalIgnoreList = Globals.ThisAddIn._ignoreList;

            //Populate ignore folder list
            /*
            _List = Globals.ThisAddIn._ignoreList;
            _List.Sort();
            

            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }
            */
            IgnoreFolders.ItemsSource = Globals.ThisAddIn._ignoreList; 
            

            //Level 1 folders in the Inbox
            _FolderLevel1L.Clear();

            if (Globals.ThisAddIn._FoldersLevel1.Count() < 1)
            {
                Globals.ThisAddIn.EnumerateFoldersInDefaultStore();
            }

            _FolderLevel1L.Clear();
            _FolderLevel1List = Globals.ThisAddIn._FoldersLevel1;
            _FolderLevel1List.Sort();

            foreach (string str in _FolderLevel1List)
            {
                if (!Globals.ThisAddIn._ignoreList.Contains(str))
                {
                    _FolderLevel1L.Add(str);
                }
            }

            InboxFolders.ItemsSource = _FolderLevel1L;
                                    
            //Update Debug checkbox
            DebugCB.IsChecked = Globals.ThisAddIn.DebugMode;

            /// Online Check Interval Drop-Down list
            TimeSpan ts = Globals.ThisAddIn._OnlineCheckInterval;
            
            UpdateCB.SelectedIndex = 0;
            for (int i = 0; i < ListData.Count; i++)
            {
                if (TimeSpan.FromMinutes(ListData[i].Id) == ts)
                {
                    UpdateCB.SelectedIndex = i;
                }
            }

            AvailableFoldersLabel.Content = "Available folders (" + InboxFolders.Items.Count + "):";
            IgnoreFoldersLabel.Content = "Ignore folders (" + IgnoreFolders.Items.Count + "):";

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

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.EnumerateFoldersInDefaultStore();
            Globals.ThisAddIn.WriteVariables();
            this.Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn._ignoreList = _originalIgnoreList;
            this.Close();
        }

        public void AddItem(object sender, EventArgs e)
        {
            List<string> _ignoreList;
            List<string> _FolderLevel1List;

            int SelectedIndex = InboxFolders.SelectedIndex;

            _ignoreList = Globals.ThisAddIn._ignoreList;
            foreach (string selected in InboxFolders.SelectedItems)
            {
                _ignoreList.Add(selected);
            }

            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
            IgnoreFolders.ItemsSource = _ignoreList;
            IgnoreFolders.Items.Refresh();


            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }

            _FolderLevel1L.Clear();
            _FolderLevel1List = Globals.ThisAddIn._FoldersLevel1;
            _FolderLevel1List.Sort();
            foreach (string str in _FolderLevel1List)
            {
                if (!Globals.ThisAddIn._ignoreList.Contains(str))
                {
                    _FolderLevel1L.Add(str);
                }
            }
            
            InboxFolders.ItemsSource = _FolderLevel1L;
            InboxFolders.Items.Refresh();

            AvailableFoldersLabel.Content = "Available folders (" + InboxFolders.Items.Count + "):";
            IgnoreFoldersLabel.Content = "Ignore folders (" + IgnoreFolders.Items.Count + "):";


            if (SelectedIndex == InboxFolders.Items.Count)
                InboxFolders.SelectedIndex = SelectedIndex - 1;
            else
                InboxFolders.SelectedIndex = SelectedIndex;            
        }

        public void RemoveItem(object sender, EventArgs e)
        {
            List<string> _FolderLevel1List;
            List<string> _ignoreList;

            int SelectedIndex = IgnoreFolders.SelectedIndex;

            _ignoreList = Globals.ThisAddIn._ignoreList;

            foreach (string selected in IgnoreFolders.SelectedItems)
            {
                _ignoreList.Remove(selected);
                _FolderLevel1L.Remove(selected);
            }

            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
            IgnoreFolders.ItemsSource = _ignoreList;
            IgnoreFolders.Items.Refresh();

            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }

            _FolderLevel1L.Clear();

            _FolderLevel1List = Globals.ThisAddIn._FoldersLevel1;
            _FolderLevel1List.Sort();
            foreach (string str in _FolderLevel1List)
            {
                if (!Globals.ThisAddIn._ignoreList.Contains(str))
                {
                    _FolderLevel1L.Add(str);
                }
            }
            
            InboxFolders.ItemsSource = _FolderLevel1L;
            InboxFolders.Items.Refresh();

            AvailableFoldersLabel.Content = "Available folders (" + InboxFolders.Items.Count + "):";
            IgnoreFoldersLabel.Content = "Ignore folders (" + IgnoreFolders.Items.Count + "):";
            

            if (SelectedIndex == InboxFolders.Items.Count)
                IgnoreFolders.SelectedIndex = SelectedIndex - 1;
            else
                IgnoreFolders.SelectedIndex = SelectedIndex;
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            AddItem(sender, e);
        }

        private void RemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            RemoveItem(sender, e);
        }

        private void DebugCB_Checked(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.DebugMode = (bool)DebugCB.IsChecked;
        }
    }
}
