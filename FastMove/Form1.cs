using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;


namespace FastMove
{
    public partial class Form1 : Form
    {
        readonly AutoCompleteStringCollection namesCollection = new AutoCompleteStringCollection();
        readonly List<string> _recentItems = new List<string>();
        readonly List<string> _items = new List<string>();
        List<string> _Searchitems = new List<string>();

        public AutoCompleteStringCollection NamesCollection => namesCollection;

        public string Pad(int i)
        {
            if (i < 10)
            {
                return "0" + i;
            }
            return i + "";
        }

        public Form1()
        {                        
            InitializeComponent();
            
            
            double seconds = Globals.ThisAddIn._InboxAvg;
            TimeSpan TS = TimeSpan.FromSeconds(seconds);                        
            string AvgText = TS.Days+ " days,"+
                Pad(TS.Hours)+ " hours, "+
                Pad(TS.Minutes)+ " minutes "+
                Pad(TS.Seconds) + " seconds";

            textBox1.Text = AvgText;

            seconds = 0;
            int count = 0;
            foreach(double d in Globals.ThisAddIn._avgTimeBeforeMove) {
                seconds += d;
                count++;
            }
            if(count>0) 
             seconds /= count;
           
            TS = TimeSpan.FromSeconds(seconds);
            AvgText = TS.Days + " days," +
                Pad(TS.Hours) + " hours, " +
                Pad(TS.Minutes) + " minutes " +
                Pad(TS.Seconds) + " seconds";

            textBox2.Text = AvgText;
   
            _items = Globals.ThisAddIn._items;
            _recentItems = Globals.ThisAddIn._recentItems;
            namesCollection = Globals.ThisAddIn.namesCollection;

            listBox1.DataSource = _items;
            listBox2.DataSource = _recentItems;

            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            comboBox1.AutoCompleteCustomSource = namesCollection;

            // Update statusbar!
            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            DateTime day = DateTime.Today.Date;
            count = 0;

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

            toolStripStatusLabel3.Text = string.Format("Version: {0}", Globals.ThisAddIn.publishedVersion);
            
            statusStrip1.Refresh();

            ///Check for updates!
            if (Globals.ThisAddIn.AddinUpdateAvailable > 0)
            {
                pictureBox1.Visible = true;
                linkLabel1.Visible = true;
            } else
            {
                pictureBox1.Visible = false;
                linkLabel1.Visible = false;
            }

        }

        private bool Compare(string s)
        {
            string t = comboBox1.Text;

            s = s.ToLower();
            t = t.ToLower();

            if (s.Contains(t))
            {
                return true;
            }
            return false;
        }

        private void ComboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length > 0)
            {
                _Searchitems = _items.FindAll(Compare);
                listBox1.DataSource = _Searchitems;
            }
        }

        private void ComboBox1_Selected(object sender, EventArgs e)
        {
            //string selected = comboBox1.SelectedText;
            object selectedItem = listBox1.SelectedItem;

            Globals.ThisAddIn.MoveMail(selectedItem.ToString());
            this.Close();
        }

        void ComboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                ComboBox1_Selected(sender, EventArgs.Empty);           
        }

        private void ComboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
                ComboBox1_Selected(sender, EventArgs.Empty);
        }        

        private void ListBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
                ComboBox1_Selected(sender, EventArgs.Empty);
        }

        private void ListBox1_MouseDoubleClick(object sender, EventArgs e)
        {            
            ComboBox1_Selected(sender, EventArgs.Empty);            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //Ok
            ComboBox1_Selected(sender, EventArgs.Empty);   
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();                
        }

        private void ListBox2_MouseDoubleClick(object sender, EventArgs e)
        {
            object selectedItem = listBox2.SelectedItem;
            Globals.ThisAddIn.MoveMail(selectedItem.ToString());
            this.Close();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                SettingsForm _Form = new SettingsForm
                {
                    Tag = this
                };
                _Form.Show();
                this.Hide();
            }
            catch (Exception ee)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: " + ee.Message);
            } 
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.EnumerateFoldersInDefaultStore();
            Globals.ThisAddIn.CalculateMeanInboxTime();
            this.Close();
        }

        private void Statistics_Click(object sender, EventArgs e)
        {
            try
            {
                StatisticsForm _Form = new StatisticsForm
                {
                    Tag = this
                };
                _Form.Show();
                this.Hide();
            }
            catch (Exception ee)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: " + ee.Message);
            }
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo("https://github.com/j-b-n/Fastmove");
            Process.Start(sInfo);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            try
            {
                DeferEmails _Form = new DeferEmails
                {
                    Tag = this
                };
                _Form.Show();
                this.Hide();
            }
            catch (Exception ee)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: " + ee.Message);
            }
        }

        private void About_Button_Click(object sender, EventArgs e)
        {
            try
            {
                About _Form = new About
                {
                    Tag = this
                };                
                _Form.Show();
                this.Hide();
            }
            catch (Exception ee)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: " + ee.Message);
            }

        }
    }
}
