using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Web;
using System.Net;
using System.IO;
using System.Reflection;
using System.Text;
using System.Diagnostics;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace FastMove
{
    public partial class ThisAddIn
    {
        #region Instance Variables

        public Office.IRibbonUI ribbon;

        public AutoCompleteStringCollection namesCollection = new AutoCompleteStringCollection();
        public List<string> _items = new List<string>();       
        public List<string> _recentItems = new List<string>();
        public List<string> _accounts = new List<string>();
        public List<string> _ignoreList = new List<string>(); 


        Microsoft.Office.Interop.Outlook.Application _applicationObject = null;
        #endregion

        #region CacheData

        public void addRecentItem(string item)
        {
            string folderStr = item;
            foreach (string str in _accounts)
            {
                folderStr = folderStr.Replace(@"\\" + str + @"\", "");
            }

            if(_recentItems.Contains(folderStr))
            {
                _recentItems.Remove(folderStr);
                _recentItems.Insert(0, folderStr);
                return;
            }
            _recentItems.Insert(0, folderStr);            
            if(_recentItems.Count > 10)
                _recentItems.RemoveAt(10); 
        }

        public void loadVariables()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove\\FastMove.xml";
            //string path = @"R:\TEMP\FastMove_recentItems.txt";                        
            //string path = @"R:\TEMP\FastMoveVariables.xml";

            try
            {
                FastMoveVariables up;

                XmlSerializer mySerializer = new XmlSerializer(typeof(FastMoveVariables));
                FileStream myFileStream = new FileStream(path, FileMode.Open);

                up = (FastMoveVariables)mySerializer.Deserialize(myFileStream);
                _recentItems = up._recentItems;
                _ignoreList = up._ignoreList;
                myFileStream.Close();

                /*
                using (StreamReader sr = new StreamReader(path))
                {
                    string line;
                    
                    while ((line = sr.ReadLine()) != null)
                    {
                        _recentItems.Add(line);
                    }

                    sr.Close();
                }
                 */
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
               MessageBox.Show("The file could not be read: "+e.Message);
            }
                        
        }

        public void writeVariables()
        {
            try
            {

            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove\\FastMove.xml";
            //string path = @"R:\TEMP\FastMove_recentItems.txt";
            //string path = @"R:\TEMP\FastMoveVariables.xml";


            FastMoveVariables up = new FastMoveVariables();

            up._ignoreList = _ignoreList;
            up._recentItems = _recentItems;
            XmlSerializer mySerializer = new XmlSerializer(typeof(FastMoveVariables));
            StreamWriter myWriter = new StreamWriter(path);
            mySerializer.Serialize(myWriter, up);
            myWriter.Close();

                /*
            int counter = 0;
            string line;
            
                // Read the file and display it line by line.
                System.IO.StreamWriter file =
                   new System.IO.StreamWriter(path);

                foreach (string item in _recentItems)
                {
                    if (counter < 11)
                        file.WriteLine(item);
                    counter++;
                }
                file.Close();
                //MessageBox.Show("Wrote: " + counter);
                 */
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The file could not be written: "+e.Message);
            }       
        }

        /// <summary>
        /// EnumerateFoldersInDefaultStore()
        /// </summary>

        public void EnumerateFoldersInDefaultStore()
        {
            _items.Clear();
            Outlook.Folder root =
                this.Application.Session.
                DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);

        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Outlook.Folder folder)
        {
            bool ignore = false;
            string folderStr = "";
            Outlook.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    ignore = false;
                    // Store the folder path. 

                    folderStr = Uri.UnescapeDataString(childFolder.FolderPath);
                    //folderStr = System.Net.WebUtility.HtmlDecode(childFolder.FolderPath);
                    //folderStr = folder.FolderPath;

                    foreach (string str in _accounts)
                    {                        
                        folderStr = folderStr.Replace(@"\\"+str+@"\","");                        
                    }

                    foreach (string str in _ignoreList)
                    {
                        if (folderStr.StartsWith(str))
                        {
                            ignore = true;
                            break;
                        }
                    }

                    if (ignore == false)
                    {
                        _items.Add(folderStr);
                        namesCollection.Add(folderStr);
                    }
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                }
            }
        }
        #endregion      

        #region MoveMail


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


        public void moveMail(string selectedFolderPath)
        {
            Outlook.MAPIFolder destFolder = null;
            bool movedMail = false;
            string itemMessage = "";

            /*Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)_addIn.Application.
              ActiveExplorer().Session.GetDefaultFolder
               (Outlook.OlDefaultFolders.olFolderInbox);*/

            //Outlook.Items items = (Outlook.Items)inBox.Items;
            //Outlook.MailItem moveMail = null;
            //items.Restrict("[UnRead] = true");

            try
            {                
                string folderStr = "";

                destFolder = GetFolder(folderStr);
                if (destFolder == null)
                {
                    foreach (string str in _accounts)
                    {
                        folderStr = @"\\" + str + @"\" + selectedFolderPath;
                        destFolder = GetFolder(folderStr);
                        if (destFolder != null)
                            break;
                    }
                }

                itemMessage += "SelFolder: " + selectedFolderPath.ToString() + "\n";

                if (destFolder != null)
                {
                    itemMessage += "DestFolder: " + destFolder.FolderPath + "\n";

                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        foreach (Object selObject in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                        {
                            if (selObject is Outlook.MailItem)
                            {
                                Outlook.MailItem mailItem =
                                   (selObject as Outlook.MailItem);
                                itemMessage += "The item is an e-mail message." +
                                    " The subject is " + mailItem.Subject + ".";
                                //mailItem.Display(false);

                                mailItem.UnRead = false;                                
                                mailItem.Move(destFolder);                              

                                itemMessage += "Moved mail!\n";
                                movedMail = true;
                                addRecentItem(Uri.UnescapeDataString(destFolder.FolderPath));
                                //MessageBox.Show(itemMessage);
                                itemMessage = "";
                            }
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
                //addRecentItem(Uri.UnescapeDataString(destFolder.FolderPath));
                //MessageBox.Show(itemMessage);                   
            }
            else
            {
                MessageBox.Show(itemMessage);
            }
        }
     

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {            
            return new Ribbon1();
        }

  
        // In case of trapping the Folder change event
        //http://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.foldersevents_event.folderchange.aspx
        /*
         *   this.Application.Session.
                DefaultStore.GetRootFolder().Folders.FolderChange +=
                new Outlook.FoldersEvents_FolderChangeEventHandler(Folders_FolderChange);            
        
*/

        /*

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            // As this is an Outlook-only extension, we know the application object will be an Outlook application 
            _applicationObject = (Microsoft.Office.Interop.Outlook.Application)application;

            // Make sure we're notified when Outlook 2010 is shutting down 
            _applicationObject.Application.Quit +=            
                new Application_QuitEventHandler(Connect_ApplicationEvents_Event_Quit);
        }

        private void Connect_ApplicationEvents_Event_Quit()
        {
            Array emptyCustomArray = new object[] { };
            OnBeginShutdown(ref emptyCustomArray);
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            addinShutdown();
        }

        public void OnBeginShutdown(ref System.Array custom)
        {
            addinShutdown();
        }

        private void addinShutdown()
        {
            // Code to run when addin is being unloaded, or Outlook is shutting down, goes here... 
            writeRecentItems();
        } 

        

        void _Explorer_FolderSwitch() 
        {
            MessageBox.Show("Folder view Switch");
        }         
        */

        
        public void HandlerQuit() {
            writeVariables();            
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
            
            foreach (Outlook.Account account in Application.Session.Accounts)
            {
                _accounts.Add(account.SmtpAddress);
            }

            loadVariables();            
            
            EnumerateFoldersInDefaultStore();

            Microsoft.Office.Interop.Outlook.Application app = this.Application;                                   

            ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)app).Quit += 
                new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_QuitEventHandler(HandlerQuit);          
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            writeVariables();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
