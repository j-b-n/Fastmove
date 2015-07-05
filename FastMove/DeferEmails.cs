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
            checkBox16.Checked = Globals.ThisAddIn._deferEmailsAlwaysSendHighPriority;

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Monday))
            {
                checkBoxMon.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Monday];
                dateTimePickerMon1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimePickerMon2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Tuesday))
            {
                checkBoxTue.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Tuesday];
                dateTimePickerTue1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimeTue2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Wednesday))
            {
                checkBoxWed.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Wednesday];
                dateTimePickerWed1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimeWed2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Thursday))
            {
                checkBoxThurs.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Thursday];
                dateTimePickerThurs1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimePickerThurs2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Friday))
            {
                checkBoxFri.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Friday];
                dateTimePickerFri1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimePickerFri2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Saturday))
            {
                checkBoxSat.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Saturday];
                dateTimePickerSat1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimePickerSat2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Sunday))
            {
                checkBoxSun.Checked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Sunday];
                dateTimePickerSun1.Value = new DateTime(2014, 01, 01) + BT.StartTS;
                dateTimePickerSun2.Value = new DateTime(2014, 01, 01) + BT.StopTS;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn._deferEmails = checkBox1.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Dictionary Weekday with list of BetweenTime
            //BetweenTime struct timespanStart and timespanStop                        

            Dictionary<DayOfWeek, BetweenTime> AllowedTime = new Dictionary<DayOfWeek, BetweenTime>();
            
            if(checkBoxMon.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerMon1.Value.TimeOfDay;
                bt.StopTS = dateTimePickerMon2.Value.TimeOfDay;
                AllowedTime.Add(DayOfWeek.Monday, bt);
            }

            if (checkBoxTue.Checked == true)
            {             
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerTue1.Value.TimeOfDay;
                bt.StopTS = dateTimeTue2.Value.TimeOfDay;             
                AllowedTime.Add(DayOfWeek.Tuesday, bt);
            }

            if (checkBoxWed.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerWed1.Value.TimeOfDay;
                bt.StopTS = dateTimeWed2.Value.TimeOfDay;             
                AllowedTime.Add(DayOfWeek.Wednesday, bt);
            }

            if (checkBoxThurs.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerThurs1.Value.TimeOfDay;
                bt.StopTS = dateTimePickerThurs2.Value.TimeOfDay;              
                AllowedTime.Add(DayOfWeek.Thursday, bt);
            }

            if (checkBoxFri.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerFri1.Value.TimeOfDay;
                bt.StopTS = dateTimePickerFri2.Value.TimeOfDay;             
                AllowedTime.Add(DayOfWeek.Friday, bt);
            }

            if (checkBoxSat.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerSat1.Value.TimeOfDay;
                bt.StopTS = dateTimePickerSat2.Value.TimeOfDay;             
                AllowedTime.Add(DayOfWeek.Saturday, bt);
            }

            if (checkBoxSun.Checked == true)
            {                
                BetweenTime bt = new BetweenTime();
                bt.StartTS = dateTimePickerSun1.Value.TimeOfDay;
                bt.StopTS = dateTimePickerSun2.Value.TimeOfDay;             
                AllowedTime.Add(DayOfWeek.Sunday, bt);
            }

            Globals.ThisAddIn._deferEmailsAllowedTime = AllowedTime;
            Globals.ThisAddIn.writeVariables();

            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();            
        }


        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn._deferEmailsAlwaysSendHighPriority = checkBox16.Checked;            
        }
    }
}
