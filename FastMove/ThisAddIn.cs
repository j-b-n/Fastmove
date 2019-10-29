using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Polenter.Serialization;
using Microsoft.Office.Interop.Outlook;
using System.Threading;

namespace FastMove
{

    public class BetweenTime
    {
        public TimeSpan StartTS {get; set;}
        public TimeSpan StopTS {get; set;}
    }

    public partial class ThisAddIn
    {
        #region Instance Variables

        public String publishedVersion = "1.0.1.7";

        public Office.IRibbonUI ribbon;

        public AutoCompleteStringCollection namesCollection = new AutoCompleteStringCollection();
        public List<string> _items = new List<string>();
        public List<string> _FoldersLevel1 = new List<string>();       

        public List<string> _recentItems = new List<string>();
        public List<string> _accounts = new List<string>();
        public List<string> _ignoreList = new List<string>();
        public double _InboxAvg = 0;
        public List<double> _avgTimeBeforeMove = new List<double>();
        
        public Dictionary<DateTime,int> _MailsPerDay = new Dictionary<DateTime, int>();
        public Dictionary<string, int> _MailsFromWho = new Dictionary<string, int>();
        public Dictionary<String, DateTime> _CountedNewMails = new Dictionary<String, DateTime>();

        public DateTime _LastMailReceived;
        public bool _LostConnection = true;

        //Defer email system
        public bool _deferEmails = false;
        public bool _deferEmailsAlwaysSendHighPriority = false;

        public Dictionary<DayOfWeek, BetweenTime> _deferEmailsAllowedTime = new Dictionary<DayOfWeek, BetweenTime>();

        /// <summary>
        /// Use Debug mode ie more logging
        /// </summary>
        public bool DebugMode = false;
        
        /// <summary>
        /// When Add-In last checked for updates
        /// </summary>
        public DateTime _LastOnlineCheck;
        
        /// <summary>
        /// Interval to check for updates online
        /// </summary>
        public TimeSpan _OnlineCheckInterval = TimeSpan.FromMinutes(60);

        private readonly System.Windows.Forms.Timer StartUpTimer = new System.Windows.Forms.Timer();
        private readonly System.Windows.Forms.Timer CalcMeanTimeTimer = new System.Windows.Forms.Timer();

        public int AddinUpdateAvailable = 0;
        private Outlook.Inspector myInspector = null;


        public DateTime _LastCalcMeanTime;

        //Microsoft.Office.Interop.Outlook.Application _applicationObject = null;
        #endregion


        #region CacheData

        public void AddRecentItem(string item)
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

        /*
         Store variables in a StorageItem?
        */

        public string GetVariables()
        {
            try
            {
                Outlook.StorageItem storage =
                    Application.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox).GetStorage(
                    "FastMove.Configuration.Variables",
                    Outlook.OlStorageIdentifierType.olIdentifyBySubject);

                Outlook.PropertyAccessor pa = storage.PropertyAccessor;
                // PropertyAccessor will return a byte array for this property

                ShowMessageBox("Vars:" + storage.Size,"");

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public string StoreVariables()
        {
            try
            {
                Outlook.StorageItem storage =
                    Application.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox).GetStorage(
                    "FastMove.Configuration.Variables",
                    Outlook.OlStorageIdentifierType.olIdentifyBySubject);

                Outlook.PropertyAccessor pa = storage.PropertyAccessor;
                // PropertyAccessor will return a byte array for this property

                pa.SetProperty("_deferEmails", _deferEmails);
                
                storage.Save();

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }


        public void LoadVariables()
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
            catch (System.Exception)
            {
                // Fail silently
            }
            path += "\\FastMove.xml";
            try
            {
                // If the directory doesn't exist, create it.
                if (!File.Exists(path))
                {
                    WriteVariables();
                }
            }
            catch (System.Exception)
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
                _LastOnlineCheck = up.LastOnlineCheck;
                _OnlineCheckInterval = up.OnlineCheckInterval;
                _CountedNewMails = up.CountedNewMails;
                _MailsFromWho = up.MailsFromWho;
                DebugMode = up.DebugMode;

                //Defer emails
                _deferEmails = up.DeferEmailActive;
                _deferEmailsAlwaysSendHighPriority = up.DeferEmailsAlwaysSendHighPriority;
                _deferEmailsAllowedTime = up.DeferEmailsAllowedTime;
            }
            catch (System.Exception e)
            {
                // Let the user know what went wrong.
                ShowMessageBox("The file could not be read: " + e.Message, "");
            }

            //GetVariables();                        
        }

        public void WriteVariables()
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
            catch (System.Exception e)
            {
                // Fail silently
                // Let the user know what went wrong.
                ShowMessageBox("WriteVariables:\n" + e.Message, "");
            }

            path += "\\FastMove.xml";

            try
            {
                PurgeCountedMails();

                FastMoveVariables up = new FastMoveVariables
                {
                    _FoldersLevel1 = _FoldersLevel1,
                    _ignoreList = _ignoreList,
                    _recentItems = _recentItems,
                    _folderItems = _items,
                    _InboxAvg = _InboxAvg,
                    _avgTimeBeforeMove = _avgTimeBeforeMove,
                    MailsPerDay = _MailsPerDay,
                    LastMailReceived = _LastMailReceived,
                    LastOnlineCheck = _LastOnlineCheck,
                    OnlineCheckInterval = _OnlineCheckInterval,
                    _CountedNewMails = _CountedNewMails,
                    MailsFromWho = _MailsFromWho,

                    DeferEmailActive = _deferEmails,
                    DeferEmailsAlwaysSendHighPriority = _deferEmailsAlwaysSendHighPriority,
                    DeferEmailsAllowedTime = _deferEmailsAllowedTime,
                    DebugMode = DebugMode
                };

                var serializer = new SharpSerializer();
                serializer.Serialize(up, path);

                /* Should I store the variables in Outlook somehow?
                 * 
                
                MemoryStream stream = new MemoryStream();
                serializer.Serialize(up, stream);
                // convert stream to string
                StreamReader reader = new StreamReader(stream);
                string text = reader.ReadToEnd();
                */

            }
            catch (System.Exception e)
            {
                // Let the user know what went wrong.
                ShowMessageBox("The file could not be written: " + e.Message,"");
            }

            //StoreVariables();
        }

        #endregion

        #region Folders

        private void ShowMessageBox(string text, string caption)
        {
            Thread t = new Thread(() => MyMessageBox(text, caption));
            t.Start();
        }

        private void MyMessageBox(object text, object caption)
        {
            MessageBox.Show((string)text, (string)caption);
        }

        public void CalcMeanTime(object sender, EventArgs e)
        {
            try
            {                
                if(_LastCalcMeanTime.AddMinutes(5) > DateTime.Now)
                {
                    //ShowMessageBox("CalcMeanTime : \nLess than 5 minutes!","");
                    return;
                }

                //Clean-up!
                CalcMeanTimeTimer.Stop();
                CalcMeanTimeTimer.Dispose();
                _LastCalcMeanTime = DateTime.Now;

                CalculateMeanInboxTime();
            }
            catch (System.Exception ex)
            {             
                ShowMessageBox("CalcMeanTime: \n"+ex.Message,"");
            }            
        }

        public void CalculateMeanInboxTime()
        {
            double avg = 0;
            int avgCount = 1;
            DateTime _now = DateTime.Now;            

            try
            {
                //Outlook.Folder folder =
                //Application.Session.GetDefaultFolder(
                // Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (DebugMode)
                    ShowMessageBox("Calc mean InboxTime: " + DateTime.Now.ToLongTimeString(), "");

                Outlook.Folder folder = (Outlook.Folder) this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                if (folder is null)
                    return; 

                foreach (Object selObject in folder.Items)
                {
                    if (selObject is Outlook.MailItem mail)
                    {
                        TimeSpan span = _now.Subtract(mail.ReceivedTime);
                        avg += span.TotalSeconds;
                        avgCount += 1;
                    }
                }

                if (avgCount > 0)
                {
                    avg /= avgCount;
                }
                else
                {
                    avg = 0;
                }
                _InboxAvg = avg;
            }
            catch (System.Exception ex)
            {                
                ShowMessageBox("CalculateMeanTome:\n"+ ex.Message, "");
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
            bool ignore;
            string folderStr;
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


        public void MoveMail(string selectedFolderPath)
        {
            Outlook.MAPIFolder destFolder;
            bool movedMail = false;
            string itemMessage = "";
            DateTime StartTime = DateTime.Now;            
            
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
                itemMessage += "Time1: " + (DateTime.Now - StartTime).TotalMilliseconds;

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
                                    " The subject is " + mailItem.Subject + ".\n";
                                //mailItem.Display(false);

                                itemMessage += "Time Count 1: " + (DateTime.Now - StartTime).TotalMilliseconds + "\n";
                                CountMail(mailItem);
                                itemMessage += "Time Count 2: " + (DateTime.Now - StartTime).TotalMilliseconds + "\n";

                                mailItem.UnRead = false;                                
                                mailItem.Move(destFolder);
                                itemMessage += "Time Move: " + (DateTime.Now - StartTime).TotalMilliseconds+"\n";

                                itemMessage += "Moved mail!\n";
                                movedMail = true;
                                AddRecentItem(Uri.UnescapeDataString(destFolder.FolderPath));

                                //ShowMessageBox(itemMessage,"");

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
            catch (System.Exception ex)
            {                
                ShowMessageBox("Movemail:\n" + ex.Message, "");
            }

            if (movedMail)
            {
                if (!CalcMeanTimeTimer.Enabled)
                {
                    CalcMeanTimeTimer.Tick += new EventHandler(CalcMeanTime);
                    CalcMeanTimeTimer.Interval = 1000 * 60 * 5; //5 minutes
                    CalcMeanTimeTimer.Enabled = true;
                    CalcMeanTimeTimer.Start();
                }                
            }
            else
            {
                ShowMessageBox("Movemail:\n" + itemMessage, "");
            }
        }
    

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {            
            return new Ribbon1();
        }


        // In case of trapping the Folder change event
        //http://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.foldersevents_event.folderchange.aspx


        public void HandlerQuit()
        {
            WriteVariables();
        }

        #region NewMail
        /*
         * Get smtp e-mailadress
         */
        public string GetEmailAdress (Outlook._MailItem MailMessage)
        {
            string SMTPAddress;

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

        void PurgeCountedMails()
        {
            object item;
            List<string> list = new List<string>();

            if (DebugMode)
                ShowMessageBox("Purge Mail: " + DateTime.Now.ToLongTimeString(), "");

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
            catch (System.Exception ex)
            {                
                ShowMessageBox("PurgeCountedMails:\n" + ex.Message, "");
            }
        }

        void CountMail(Outlook.MailItem item)
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
            if (DebugMode)
                ShowMessageBox("NewMailEx: " + DateTime.Now.ToLongTimeString(), "");
            CountMail(((Outlook.MailItem)this.Application.Session.GetItemFromID(entryIDCollection, missing)));
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
                if (DebugMode)
                    ShowMessageBox("ItemsAdd: " + DateTime.Now.ToLongTimeString(), "");
                CountMail((Outlook.MailItem)item);
            }            
        }

        public void Items_ItemRemove(object item)
        {
            if (item is Outlook.MailItem)
            {
                if (DebugMode)
                    ShowMessageBox("ItemsRemove: " + DateTime.Now.ToLongTimeString(), "");
                CountMail((Outlook.MailItem)item);
            }
        }

        #region DeferEmails
        public static DateTime GetNextWeekday(DateTime start, DayOfWeek day)
        {
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            int daysToAdd = ((int)day - (int)start.DayOfWeek + 7) % 7;
            return start.AddDays(daysToAdd);
        }

        /// <summary>
        /// Determines if a mail can be sent!
        /// </summary>
        /// <param name="sendTime"></param>
        /// <returns></returns>
        public bool AllowedToSendDirectly(DateTime sendTime)
        {
            if (_deferEmailsAllowedTime.ContainsKey(sendTime.DayOfWeek))
            {
                BetweenTime BT = _deferEmailsAllowedTime[sendTime.DayOfWeek];
             
                if((sendTime.Date+BT.StartTS).CompareTo(sendTime)<=0 &&
                    (sendTime.Date + BT.StopTS).CompareTo(sendTime) >= 0)
                {
                    return true;
                }

                return false;
            }
            return false;
        }

        public DateTime NextPossibleSendTime()
        {
            DateTime Next = DateTime.Now;
            bool _found = false;
            
            while (_found == false)
            {
                if (_deferEmailsAllowedTime.ContainsKey(Next.DayOfWeek))
                {
                    BetweenTime BT = _deferEmailsAllowedTime[Next.DayOfWeek];

                    DateTime DT = Next.Date + BT.StartTS;                                     
                    if (DT > DateTime.Now && AllowedToSendDirectly(DT))
                    {
                        return DT;
                    }
                }

                Next = Next.AddDays(1);

                if(Next > DateTime.Now.AddDays(7))
                {
                    return DateTime.Now.AddMinutes(10);
                }

                
            }          
            return Next;
        }

        private void DeferEmail(object Item, ref bool Cancel)
        {
            var msg = Item as Outlook.MailItem;
            //DateTime sendTime;
            DateTime deferTime;

            if (_deferEmails == false)
            {                
                return;
            }

            if(msg.Importance == Outlook.OlImportance.olImportanceHigh &&
                _deferEmailsAlwaysSendHighPriority == true)
            {                
                return;
            }            

            if(AllowedToSendDirectly(DateTime.Now))
            {             
                return;
            }
           
            deferTime = NextPossibleSendTime();
            
            AutoClosingMessageBox.Show("Sending mail at: "+ deferTime.ToString(),"Defer time",600);
            
            msg.DeferredDeliveryTime = deferTime;
        }

        #endregion


        /* Below are some checks to see if the addin is online or not!
         * 
         *
        void Timer_Tick(object sender, EventArgs e)
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
                _LostConnection = true;
            }
            _LostConnection = false;
        }
        */

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            myInspector = Inspector;
            if (myInspector.CurrentItem is Outlook.MailItem)
            {
                if (DebugMode)
                    ShowMessageBox("New mail: " + DateTime.Now.ToLongTimeString(), "");
                CountMail(myInspector.CurrentItem);                
            }
        }

        #endregion



        /// <summary>
        /// Sometimes the initialization takes time and causes Outlook to disable the plugin. 
        /// This is solved using a delayed startup sequence.
        /// </summary>
        void DelayedStartup(object sender, EventArgs e)
        {
            if (_items.Count < 1)
                EnumerateFoldersInDefaultStore();

            CalculateMeanInboxTime();            

            if (_LastOnlineCheck.Add(_OnlineCheckInterval) < DateTime.Now)
            {
                AddinUpdateAvailable = (new UpdateInfo()).CheckForUpdate();
                _LastOnlineCheck = DateTime.Now;
            }

            //Count mails in inbox
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.
               ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items items = (Outlook.Items)inBox.Items;
            items = items.Restrict("[Unread] = true");
            try
            {
                if (items != null && items.Count > 0)
                {
                    foreach (object item in items)
                    {
                        if (item is Outlook.MailItem obj)
                        {
                            CountMail(obj);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ShowMessageBox("DelayedStartup: " + ex.Message, "");
            }

            //Clean-up!
            StartUpTimer.Stop();
            StartUpTimer.Dispose();
        }

        /// <summary>
        /// ThisAddIn is called at startup of the plugin!
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            foreach (Outlook.Account account in Application.Session.Accounts)
            {
                _accounts.Add(account.SmtpAddress);
            }

            LoadVariables();

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
           
            //Create an event handler for when items are sent
            Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(DeferEmail);


            /*
             this.Application.Session.
                  DefaultStore.GetRootFolder().Folders.FolderAdd +=
                  new Outlook.FoldersEvents_FolderAddEventHandler(Folders_FolderAdd);

             this.Application.Session.
                  DefaultStore.GetRootFolder().Folders.FolderRemove +=
                  new Outlook.FoldersEvents_FolderRemoveEventHandler(Folders_FolderRemove);            
         */

            StartUpTimer.Tick += new EventHandler(DelayedStartup);
            StartUpTimer.Interval = 1000 * 5; //5 seconds
            StartUpTimer.Enabled = true;
            StartUpTimer.Start();                              
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            WriteVariables();
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
