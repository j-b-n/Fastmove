using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FastMove
{
    public class FastMoveVariables
    {
        internal List<string> _ignoreList = new List<string>();                 
        public List<string> IgnoreList
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

        public List<string> RecentItems
        {
            get { return _recentItems; }
            set { _recentItems = value; }
        }

        internal List<string> _folderItems = new List<string>();

        public List<string> FolderItems
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

        public List<double> AvgTimeBeforeMove
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

        public DateTime _LastOnlineCheck;
        public DateTime LastOnlineCheck
        {
            get { return _LastOnlineCheck; }
            set { _LastOnlineCheck = value; }
        }

        public TimeSpan _OnlineCheckInterval;
        public TimeSpan OnlineCheckInterval
        {
            get { return _OnlineCheckInterval; }
            set { _OnlineCheckInterval = value; }
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

        //Defer email variables

        public bool _DeferEmailActive;

        public bool DeferEmailActive
        {
            get { return _DeferEmailActive; }
            set { _DeferEmailActive = value; }
        }

        public bool _deferEmailsAlwaysSendHighPriority;
        public bool DeferEmailsAlwaysSendHighPriority
        {
            get { return _deferEmailsAlwaysSendHighPriority; }
            set { _deferEmailsAlwaysSendHighPriority = value; }
        }

        public Dictionary<DayOfWeek, BetweenTime> _deferEmailsAllowedTime = new Dictionary<DayOfWeek, BetweenTime>();
        public Dictionary<DayOfWeek, BetweenTime> DeferEmailsAllowedTime
        {
            get { return _deferEmailsAllowedTime; }
            set { _deferEmailsAllowedTime = value; }
        }
    }
}
