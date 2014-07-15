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
        private BindingList<string> _ignoreL = new BindingList<string>();
        List<string> _originalIgnoreList = new List<string>(); 

        public SettingsForm()
        {
            InitializeComponent();
        }
        
        private void SettingsForm_Load(object sender, EventArgs e)
        {   
            List<string> _ignoreList = new List<string>(); 

            _ignoreList = Globals.ThisAddIn._ignoreList;
            _ignoreList.Sort();
            _originalIgnoreList = _ignoreList;

            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }
            
            listBox1.DataSource = _ignoreL;            
        }

        void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, EventArgs.Empty);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string txt = textBox1.Text;
            List<string> _ignoreList = new List<string>();

            _ignoreList = Globals.ThisAddIn._ignoreList;
            _ignoreList.Add(txt);
            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }
            
            textBox1.Clear();            
        }


        private void button2_Click(object sender, EventArgs e)
        {
            List<string> _ignoreList = new List<string>();
            _ignoreList = Globals.ThisAddIn._ignoreList;
            foreach (string selected in listBox1.SelectedItems)
            {
                _ignoreList.Remove(selected);                
            }
                        
            _ignoreList.Sort();
            Globals.ThisAddIn._ignoreList = _ignoreList;
            _ignoreL.Clear();
            foreach (string str in Globals.ThisAddIn._ignoreList)
            {
                _ignoreL.Add(str);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.EnumerateFoldersInDefaultStore();
            Globals.ThisAddIn.writeVariables();
            this.Close();
        }

        private void cancel_Click(object sender, EventArgs e)
        {            
            Globals.ThisAddIn._ignoreList = _originalIgnoreList;
            this.Close();
        }

    }
}
