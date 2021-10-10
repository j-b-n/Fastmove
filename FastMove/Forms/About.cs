using System;
using System.Windows.Forms;

namespace FastMove
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void OK_button_Click(object sender, EventArgs e)
        {
            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();
        }

        private void About_Load(object sender, EventArgs e)
        {
            UpdateInfo ui = new UpdateInfo();
            
            int AddinUpdateAvailable = ui.CheckForUpdate();
            string runningVersion = Globals.ThisAddIn.publishedVersion;
           
            this_label.Text = string.Format("Version: '{0}'", runningVersion);
            online_label.Text = string.Format("Online version: '{0}'", ui.UpdateVariables.Version);
        }
    }
}
