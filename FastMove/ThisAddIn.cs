﻿using System;
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
using Polenter.Serialization;

namespace FastMove
{
    public partial class ThisAddIn
    {
        #region Instance Variables

        public Office.IRibbonUI ribbon;

        public AutoCompleteStringCollection namesCollection = new AutoCompleteStringCollection();
        public List<string> _items = new List<string>();
        public List<string> _FoldersLevel1 = new List<string>();       

        public List<string> _recentItems = new List<string>();
        public List<string> _accounts = new List<string>();
        public List<string> _ignoreList = new List<string>();
        public double _InboxAvg = 0;
        public List<double> _avgTimeBeforeMove = new List<double>();
        public String publishedVersion = "0:0:0:-2";
        public Dictionary<DateTime,int> _MailsPerDay = new Dictionary<DateTime, int>();
        public Dictionary<string, int> _MailsFromWho = new Dictionary<string, int>();
        public Dictionary<String, DateTime> _CountedNewMails = new Dictionary<String, DateTime>();

        public DateTime _LastMailReceived;
        public bool _LostConnection = true;

        Timer timer = new Timer();
        int timerCounter = 0;

        private Outlook.Inspector myInspector = null;


        //Microsoft.Office.Interop.Outlook.Application _applicationObject = null;
        #endregion

        #region CurrentVersion

        private void GetRunningVersion()
        {
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                System.Deployment.Application.ApplicationDeployment currDeploy = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
                Version pubVer = currDeploy.CurrentVersion;
                publishedVersion = pubVer.Major.ToString() + "." + pubVer.Minor.ToString() + "." + 
                    pubVer.Build.ToString() + "." + pubVer.Revision.ToString();
                return;
            }
            publishedVersion = "0:0:0:-1";
        }
        
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
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove";

            try
            {
                // If the directory doesn't exist, create it.
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch (Exception)
            {
                // Fail silently
            }
            path += "\\FastMove.xml";
            try
            {
                // If the directory doesn't exist, create it.
                if (!File.Exists(path))
                {
                    writeVariables();
                }
            }
            catch (Exception)
            {
                // Fail silently
            }

            try
            {
                FastMoveVariables up;
                var serializer = new SharpSerializer();
                up = (FastMoveVariables)serializer.Deserialize(path);

                _FoldersLevel1 = up._FoldersLevel1;
                _recentItems = up._recentItems;
                _items = up._folderItems;
                _ignoreList = up._ignoreList;
                _InboxAvg = up._InboxAvg;
                _avgTimeBeforeMove = up._avgTimeBeforeMove;
                _MailsPerDay = up.MailsPerDay;
                _LastMailReceived = up.LastMailReceived;
                _CountedNewMails = up.CountedNewMails;
                _MailsFromWho = up.MailsFromWho;
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
               MessageBox.Show("The file could not be read: "+e.Message);
            }
                        
        }

        public void writeVariables()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove";

            try
            {
                // If the directory doesn't exist, create it.
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch (Exception)
            {
                // Fail silently
            }

            path += "\\FastMove.xml";

            try
            {
                purgeCountedMails();

                FastMoveVariables up = new FastMoveVariables();

                up._FoldersLevel1 = _FoldersLevel1;
                up._ignoreList = _ignoreList;
                up._recentItems = _recentItems;
                up._folderItems = _items;
                up._InboxAvg = _InboxAvg;
                up._avgTimeBeforeMove = _avgTimeBeforeMove;
                up.MailsPerDay = _MailsPerDay;
                up.LastMailReceived = _LastMailReceived;
                up._CountedNewMails = _CountedNewMails;
                up.MailsFromWho = _MailsFromWho;

                var serializer = new SharpSerializer();
                serializer.Serialize(up, path);
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The file could not be written: " + e.Message);
            }
        }

        #endregion
        
        #region Folders
        
        public void CalculateMeanInboxTime()
        {
            double avg = 0;
            int avgCount = 1;            
            DateTime _now = DateTime.Now;

             try
            {

            Outlook.Folder folder =
             Application.Session.GetDefaultFolder(
              Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            
            foreach (Object selObject in folder.Items)
            {
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mail = (Outlook.MailItem)selObject;
                    TimeSpan span = _now.Subtract(mail.ReceivedTime);
                    avg += span.TotalSeconds;
                    avgCount += 1;
                }
            }

            if (avgCount > 0)
            {
                avg = avg / avgCount;
            }
            else
            {
                avg = 0;
            }
            _InboxAvg = avg;
              }
             catch (Exception ex)
             {
                 string expMessage = ex.Message;
                 MessageBox.Show(expMessage);
             }


        }

        /// <summary>
        /// EnumerateFoldersInDefaultStore()
        /// </summary>

        public void EnumerateFoldersInDefaultStore()
        {
            _FoldersLevel1.Clear();
            _items.Clear();

            Outlook.Folder root =
                this.Application.Session.
                DefaultStore.GetRootFolder() as Outlook.Folder;

            EnumerateFolders(root, 0);
        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Outlook.Folder folder, int level)
        {
            bool ignore = false;
            string folderStr = "";
            Outlook.Folders childFolders = folder.Folders;
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
                        folderStr = folderStr.Replace(@"\\" + str + @"\", "");
                    }

                    foreach (string str in _ignoreList)
                    {
                        if (folderStr.StartsWith(str))
                        {
                            ignore = true;
                            break;
                        }
                    }

                    List<string> myList = new List<string>(folderStr.Split('\\'));
                    if (myList.Count > 0)
                    {
                        if (!_FoldersLevel1.Contains(myList[0]))
                        {
                            _FoldersLevel1.Add(myList[0]);
                        }
                    }                    

                    if (ignore == false)
                    {
                        _items.Add(folderStr);
                        namesCollection.Add(folderStr);
                    }
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder, level++);
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
                if (selectedFolderPath.Length < 1)
                {
                    return;
                }

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

                                countMail(mailItem);

                                mailItem.UnRead = false;                                
                                mailItem.Move(destFolder);                              

                                itemMessage += "Moved mail!\n";
                                movedMail = true;
                                addRecentItem(Uri.UnescapeDataString(destFolder.FolderPath));
                                //MessageBox.Show(itemMessage);
                                itemMessage = "";

                                DateTime _now = DateTime.Now;
                                TimeSpan span = _now.Subtract(mailItem.ReceivedTime);
                                _avgTimeBeforeMove.Add(span.TotalSeconds);
                                if (_avgTimeBeforeMove.Count > 100)
                                    _avgTimeBeforeMove.RemoveAt(0); 
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
                CalculateMeanInboxTime();
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
        
        
        public void Folders_FolderAdd(object Folder)
        {
        }

        public void Folders_FolderRemove(object Folder)
        {
        }

        public void HandlerQuit()
        {
            writeVariables();
        }

        #region NewMail
        /*
         * Get smtp e-mailadress
         */
        public string GetEmailAdress (Outlook._MailItem MailMessage)
        {
            string SMTPAddress =  "";

            //Issue a reply on the mail message to create a recipient object that is the sender address.

            if (MailMessage.SenderEmailType ==  "EX")
            {
                Outlook._MailItem Temp = ((Outlook._MailItem)MailMessage).Reply();

                //Use the recipient object to access the smtp address of the exchange user
                SMTPAddress = Temp.Recipients[1].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                Temp.Delete();
            }
            else
            {
                SMTPAddress = MailMessage.SenderEmailAddress;
            }
            return SMTPAddress;
        }

        /*
         * void purgeCountedMails()
         * 
         * Deletes all non-present mailitems in the Dictionary _CountedNewMails
         * 
         */

        public object GetObjectBasedOnEntryID(string id)
        {
            try
            {
                object item = (Outlook.MailItem)this.Application.Session.GetItemFromID(id);
                return item;
            }
            catch
            {
                return null;
            }
        }

        void purgeCountedMails()
        {
            object item = null;
            List<string> list = new List<string>(); 
            
            try
            {                
                foreach (string id in _CountedNewMails.Keys)
                {
                    list.Add(id);                    
                }

                foreach(string id in list) 
                {
                    item = GetObjectBasedOnEntryID(id);

                    if (item == null)
                    {
                        _CountedNewMails.Remove(id);
                    }
                }

            }
            catch (Exception ex)
            {
                string itemMessage = ex.Message;
                MessageBox.Show(itemMessage);
            }
        }

        void countMail(Outlook.MailItem item)
        {
            if (item == null)
            {
                return;
            }

            if (_CountedNewMails.ContainsKey(item.EntryID))
            {
                return; //The mail has been counted!
            }

            if (item.ReceivedTime > _LastMailReceived)
            {
                _LastMailReceived = item.ReceivedTime;
            }
            

            if (_MailsPerDay.ContainsKey(item.ReceivedTime.Date))
            {
                _MailsPerDay[item.ReceivedTime.Date] = _MailsPerDay[item.ReceivedTime.Date] + 1;
            }
            else
            {
                _MailsPerDay.Add(item.ReceivedTime.Date, 1);
            }


            string fromWho = item.SenderName + " (" + GetEmailAdress(item) + ")";
            if (_MailsFromWho.ContainsKey(fromWho))
            {
                _MailsFromWho[fromWho] = _MailsFromWho[fromWho] + 1;
            }
            else
            {
                _MailsFromWho.Add(fromWho, 1);

            }            

            _CountedNewMails.Add(item.EntryID, item.ReceivedTime);            
        }

        void HandlerNewMailEx(string entryIDCollection)
        {
            countMail(((Outlook.MailItem)this.Application.Session.GetItemFromID(entryIDCollection, missing)));
        }
                
        /*
        public void HandlerNewMail()
        {
                string itemMessage = " New mail event! ";
                MessageBox.Show(itemMessage);                                
        }
         * */

        public void Items_ItemAdd(object item)
        {
            if (item is Outlook.MailItem)
            {
                countMail((Outlook.MailItem)item);
            }            
        }

        public void Items_ItemRemove(object item)
        {
            if (item is Outlook.MailItem)
            {
                countMail((Outlook.MailItem)item);
            }
        }

        void timer_Tick(object sender, EventArgs e)
        {
            if (Application.Session.ExchangeConnectionMode == Outlook.OlExchangeConnectionMode.olCachedDisconnected
                ||
                Application.Session.ExchangeConnectionMode == Outlook.OlExchangeConnectionMode.olDisconnected
                ||
                Application.Session.ExchangeConnectionMode == Outlook.OlExchangeConnectionMode.olNoExchange
                ||
                Application.Session.ExchangeConnectionMode == Outlook.OlExchangeConnectionMode.olOffline
                ||
                Application.Session.ExchangeConnectionMode == Outlook.OlExchangeConnectionMode.olCachedOffline
                )
            {
                if (_LostConnection == false)
                {
                    timer.Stop();                              // Stop the timer                
                    timer.Interval = (1000) * (5);              // Timer will tick X second
                    timer.Enabled = true;                       // Enable the timer                
                    timer.Start();
                }
                _LostConnection = true;
                return;
            }

            if (_LostConnection == true)
            {
                timer.Stop();                              // Stop the timer                
                timer.Interval = (1000) * (60);            // Timer will tick X second
                timer.Enabled = true;                      // Enable the timer                
                timer.Start();
                timerCounter = 0;
            }
              

            _LostConnection = false;

            if (timerCounter < 5)
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.
                  ActiveExplorer().Session.GetDefaultFolder
                   (Outlook.OlDefaultFolders.olFolderInbox);

                Outlook.Items items = (Outlook.Items)inBox.Items;

                foreach (Outlook.MailItem item in items)
                {
                    if (item is Outlook.MailItem)
                    {
                        countMail((Outlook.MailItem)item);
                    }
                }
            }

            timerCounter++;                       
        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            myInspector = Inspector;
            if (myInspector.CurrentItem is Outlook.MailItem)
            {
                countMail(myInspector.CurrentItem);
                //MessageBox.Show("Inspector!");
            }
        }


        #endregion

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            foreach (Outlook.Account account in Application.Session.Accounts)
            {
                _accounts.Add(account.SmtpAddress);
            }

            loadVariables();            
            CalculateMeanInboxTime();
            GetRunningVersion();
            if (_items.Count < 1)
                EnumerateFoldersInDefaultStore();

            Microsoft.Office.Interop.Outlook.Application app = this.Application;

            ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)app).Quit +=
                new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_QuitEventHandler(HandlerQuit);

            /*
            ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)app).NewMail +=
                new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(HandlerNewMail);
            */

            ((Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)app).NewMailEx +=
                new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailExEventHandler(HandlerNewMailEx);

            Outlook.Folder folder =
             Application.Session.GetDefaultFolder(
              Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;

            folder.Items.ItemAdd += new Microsoft.Office.Interop.Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);                       
            
            Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            timer.Tick += new EventHandler(timer_Tick); // Everytime timer ticks, timer_Tick will be called
            timer.Interval = (1000) * (60);              // Timer will tick X second
            timer.Enabled = true;                       // Enable the timer
            timer.Start();                              // Start the timer



            /*
             this.Application.Session.
                  DefaultStore.GetRootFolder().Folders.FolderAdd +=
                  new Outlook.FoldersEvents_FolderAddEventHandler(Folders_FolderAdd);

             this.Application.Session.
                  DefaultStore.GetRootFolder().Folders.FolderRemove +=
                  new Outlook.FoldersEvents_FolderRemoveEventHandler(Folders_FolderRemove);            
         */

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
