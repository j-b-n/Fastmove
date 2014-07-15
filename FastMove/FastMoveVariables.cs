using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FastMove
{
    public class FastMoveVariables
    {
        internal List<string> _ignoreList = new List<string>();                 
        public List<string> ignoreList
        {
            get { return _ignoreList; }
            set { _ignoreList = value; }
        }

        internal List<string> _FoldersLevel1 = new List<string>();
        public List<string> FoldersLevel1
        {
            get { return _FoldersLevel1; }
            set { _FoldersLevel1 = value; }
        }

        internal List<string> _recentItems = new List<string>();

        public List<string> recentItems
        {
            get { return _recentItems; }
            set { _recentItems = value; }
        }

        internal List<string> _folderItems = new List<string>();

        public List<string> folderItems
        {
            get { return _folderItems; }
            set { _folderItems = value; }
        }

        internal double _InboxAvg = 0;
        public double InboxAvg
        {
            get { return _InboxAvg; }
            set { _InboxAvg = value; }
        }

        
        internal List<double> _avgTimeBeforeMove = new List<double>();

        public List<double> avgTimeBeforeMove
        {
            get { return _avgTimeBeforeMove; }
            set { _avgTimeBeforeMove = value; }
        }

        public DateTime _LastMailReceived;
        public DateTime LastMailReceived 
        {
            get { return _LastMailReceived; }
            set { _LastMailReceived = value; }
        }

        public Dictionary<DateTime, int> _MailsPerDay = new Dictionary<DateTime, int>();
        public Dictionary<DateTime, int> MailsPerDay
        {
            get { return _MailsPerDay; }
            set { _MailsPerDay = value; }        
        }


        public Dictionary<string, int> _MailsFromWho = new Dictionary<string, int>();
        public Dictionary<string, int> MailsFromWho
        {
            get { return _MailsFromWho; }
            set { _MailsFromWho = value; }        
        }


        public Dictionary<String, DateTime> _CountedNewMails = new Dictionary<String, DateTime>();
        public Dictionary<String, DateTime> CountedNewMails
        {
            get { return _CountedNewMails; }
            set { _CountedNewMails = value; }
        }

    }
}
