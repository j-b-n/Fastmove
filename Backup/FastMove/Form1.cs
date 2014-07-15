using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Web;
using System.Net;
using System.IO;


namespace FastMove
{
    public partial class Form1 : Form
    {        
        AutoCompleteStringCollection namesCollection = new AutoCompleteStringCollection();
        List<string> _recentItems = new List<string>(); 
        List<string> _items = new List<string>();
        List<string> _Searchitems = new List<string>(); 
        

        /*
        // Returns Folder object based on folder path
        private Outlook.Folder GetFolder(string folderPath)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders =
                    folderPath.Split(backslash.ToCharArray());
                folder =
                    Globals.ThisAddIn.Application.Session.Folders[folders[0]]
                    as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]]
                            as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch { return null; }            
        }        

        


        public void moveMail(object selectedFolder)
        {
            Outlook.MAPIFolder destFolder = null;
            bool movedMail = false;
            string itemMessage = "";

            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)_addIn.Application.
              ActiveExplorer().Session.GetDefaultFolder
               (Outlook.OlDefaultFolders.olFolderInbox);
            
            //Outlook.Items items = (Outlook.Items)inBox.Items;
            //Outlook.MailItem moveMail = null;
            //items.Restrict("[UnRead] = true");

            try
            {               
                destFolder = GetFolder(selectedFolder.ToString());
                itemMessage += "SelFolder: " + selectedFolder.ToString() + "\n";

                if (destFolder != null)
                {
                    itemMessage += "DestFolder: " + destFolder.FolderPath + "\n";

                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        Object selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                        if (selObject is Outlook.MailItem)
                        {
                            Outlook.MailItem mailItem =
                                (selObject as Outlook.MailItem);
                            itemMessage += "The item is an e-mail message." +
                                " The subject is " + mailItem.Subject + ".";
                            //mailItem.Display(false);
                          
                            mailItem.Move(destFolder);
                            itemMessage += "Moved mail!\n";
                            movedMail = true;
                        }
                    }
                }
                else
                {
                    itemMessage += "DestFolder: NULL\n";
                }
                                                 
            }
            catch (Exception ex)
            {
                itemMessage = ex.Message;
                MessageBox.Show(itemMessage);
            }

            if (movedMail)
            {
                Globals.ThisAddIn.addRecentItem(Uri.UnescapeDataString(destFolder.FolderPath));
                //MessageBox.Show(itemMessage);
                this.Close();
            }
            else
            {
                MessageBox.Show(itemMessage);
            }
        }
     */
        public Form1()
        {                        
            InitializeComponent();
            
            _items = Globals.ThisAddIn._items;
            _recentItems = Globals.ThisAddIn._recentItems;
            namesCollection = Globals.ThisAddIn.namesCollection;

            listBox1.DataSource = _items;
            listBox2.DataSource = _recentItems;

            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            comboBox1.AutoCompleteCustomSource = namesCollection;                    
        }

        private bool compare(string s)
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

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length > 0)
            {
                _Searchitems = _items.FindAll(compare);
                listBox1.DataSource = _Searchitems;
            }
        }

        private void comboBox1_Selected(object sender, EventArgs e)
        {
            string selected = comboBox1.SelectedText;
            object selectedItem = listBox1.SelectedItem;

            Globals.ThisAddIn.moveMail(selectedItem.ToString());
            this.Close();
        }

        void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                comboBox1_Selected(sender, EventArgs.Empty);           
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
                comboBox1_Selected(sender, EventArgs.Empty);
        }        

        private void listBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
                comboBox1_Selected(sender, EventArgs.Empty);
        }

        private void listBox1_MouseDoubleClick(object sender, EventArgs e)
        {            
            comboBox1_Selected(sender, EventArgs.Empty);            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Ok
            comboBox1_Selected(sender, EventArgs.Empty);   
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();                
        }

        private void listBox2_MouseDoubleClick(object sender, EventArgs e)
        {
            object selectedItem = listBox2.SelectedItem;
            Globals.ThisAddIn.moveMail(selectedItem.ToString());
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {                
                SettingsForm _Form = new SettingsForm();
                _Form.Show();
                this.Close();
            }
            catch (Exception ee)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: " + ee.Message);
            } 
        }        
    }
}
