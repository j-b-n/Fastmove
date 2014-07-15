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

        internal List<string> _recentItems = new List<string>();

        public List<string> recentItems
        {
            get { return _recentItems; }
            set { _recentItems = value; }
        }

    }
}
