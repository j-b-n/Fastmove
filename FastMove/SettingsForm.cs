using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FastMove
{
    public partial class SettingsForm : Form
    {
        /// <summary>
        /// Intervals for online check 1 hr, 2 hrs, 1 day, 1 week
        /// </summary>
        private readonly int[] OnlineCheckIntervalValues = new int[] { 60, 120, 1440, 10080 };

        private readonly BindingList<string> _ignoreL = new BindingList<string>();
        private readonly BindingList<string> _FolderLevel1L = new BindingList<string>();
        List<string> _originalIgnoreList = new List<string>(); 

        public SettingsForm()
        {
            InitializeComponent();
        }
        
        private void SettingsForm_Load(object sender, EventArgs e)
        {   
            List<string> _List ;
            List<string> _FolderLevel1List;
            
            _List = Globals.ThisAddIn._ignoreList;
            _List.Sort();
            _originalIgnoreList = _List;

            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }

            listBox1.DataSource = null;
            listBox1.DataSource = _ignoreL;

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
            listBox2.DataSource = null;
            listBox2.DataSource = _FolderLevel1L;
                                

            // Update statusbar!
            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            DateTime day = DateTime.Today.Date;
            int count = 0;

            if (_MailsPerDay.ContainsKey(day))
            {
                count = _MailsPerDay[day];
            }
            toolStripStatusLabel1.Text = string.Format("Today: {0}", count);

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

            toolStripStatusLabel2.Text = string.Format("Last week: {0}", count);
            toolStripStatusLabel4.Text = string.Format("Version: {0}", Globals.ThisAddIn.publishedVersion);
            statusStrip1.Refresh();

            /// Online Check Interval Drop-Down list
            TimeSpan ts = Globals.ThisAddIn._OnlineCheckInterval;
            comboBox1.SelectedIndex = 0;
            for (int i = 0; i < OnlineCheckIntervalValues.Length; i++ )
            {
                if (TimeSpan.FromMinutes(OnlineCheckIntervalValues[i]) == ts)
                {
                    comboBox1.SelectedIndex = i;                    
                }
            }            
        }

        public void AddItem(object sender, EventArgs e)
        {
            List<string> _ignoreList;
            List<string> _FolderLevel1List;


            _ignoreList = Globals.ThisAddIn._ignoreList;
            foreach (string selected in listBox2.SelectedItems)
            {
                _ignoreList.Add(selected);
            }

            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
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
            listBox2.DataSource = null;
            listBox2.DataSource = _FolderLevel1L;

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            AddItem(sender, e);
        }

        public void RemoveItem(object sender, EventArgs e)        
        {
            List<string> _FolderLevel1List;
            List<string> _ignoreList;
            _ignoreList = Globals.ThisAddIn._ignoreList;
            foreach (string selected in listBox1.SelectedItems)
            {
                _ignoreList.Remove(selected);
                _FolderLevel1L.Remove(selected);
            }

            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
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
            listBox2.DataSource = null;
            listBox2.DataSource = _FolderLevel1L;
        }


        private void Button2_Click(object sender, EventArgs e)
        {
            RemoveItem(sender, e);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.EnumerateFoldersInDefaultStore();
            Globals.ThisAddIn.WriteVariables();
            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();            
        }

        private void Cancel_Click(object sender, EventArgs e)
        {            
            Globals.ThisAddIn._ignoreList = _originalIgnoreList;
            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {            
            int selectedIndex = comboBox1.SelectedIndex;
            if (selectedIndex < OnlineCheckIntervalValues.Length)
            {
                Globals.ThisAddIn._OnlineCheckInterval = TimeSpan.FromMinutes(OnlineCheckIntervalValues[selectedIndex]);
            } 
        }
       
    }
}
