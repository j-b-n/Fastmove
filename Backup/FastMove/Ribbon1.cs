using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace FastMove
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {        
        private Office.IRibbonUI ribbon;
        List<string> lastList = new List<string>(); // <-- Add this


        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FastMove.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public string menu1_GetContent(Office.IRibbonControl control)
        {
            string tmp = "";
            int count = 0;
            StringBuilder MyStringBuilder = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");
            foreach (string item in Globals.ThisAddIn._recentItems)
            {
                tmp = item;
                foreach (string str in Globals.ThisAddIn._accounts)
                {
                    tmp = tmp.Replace(@"\\" + str + @"\", "");
                }
                MyStringBuilder.Append("<button id=\"btn"+count+
                    "\" imageMso=\"MoveItem\" label=\"" + tmp +
                    "\" onAction=\"FastMoveMail_btn"+count+"\"/>");
                count++;
            }
            MyStringBuilder.Append(@"</menu>");            
            return MyStringBuilder.ToString();            
        }

        public void FastMoveMail_btn0(Office.IRibbonControl control)
        {
            if(Globals.ThisAddIn._recentItems.Count >= 0) 
             Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[0].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn1(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 1)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[1].ToString());
            this.ribbon.Invalidate();
        }
        
        public void FastMoveMail_btn2(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 2)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[2].ToString());
            this.ribbon.Invalidate();
        }
        
        public void FastMoveMail_btn3(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 3)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[3].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn4(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 4)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[4].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn5(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 5)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[5].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn6(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 6)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[6].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn7(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 7)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[7].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn8(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 8)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[8].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail_btn9(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn._recentItems.Count >= 9)
                Globals.ThisAddIn.moveMail(Globals.ThisAddIn._recentItems[9].ToString());
            this.ribbon.Invalidate();
        }

        public void FastMoveMail(Office.IRibbonControl control)
        {
            try
            {
                this.ribbon.Invalidate();
                Form1 _Form = new Form1();
                _Form.Show();
                
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
                MessageBox.Show("The form could not be loaded: "+e.Message);
            }            

        }        

       #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
