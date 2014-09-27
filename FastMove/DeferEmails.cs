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
    public partial class DeferEmails : Form
    {
        public DeferEmails()
        {
            InitializeComponent();
            checkBox1.Checked = Globals.ThisAddIn._deferEmails;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn._deferEmails = checkBox1.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.writeVariables();
            this.Close();
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
