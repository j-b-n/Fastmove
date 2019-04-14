using System;
using System.Collections.Generic;
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

            DateTime next = Globals.ThisAddIn.NextPossibleSendTime();
            label15.Text = "The next possible send time is: "+next.ToString();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn._deferEmails = checkBox1.Checked;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //Dictionary Weekday with list of BetweenTime
            //BetweenTime struct timespanStart and timespanStop                        

            Dictionary<DayOfWeek, BetweenTime> AllowedTime = new Dictionary<DayOfWeek, BetweenTime>();
            
            if(checkBoxMon.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerMon1.Value.TimeOfDay,
                    StopTS = dateTimePickerMon2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Monday, bt);
            }

            if (checkBoxTue.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerTue1.Value.TimeOfDay,
                    StopTS = dateTimeTue2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Tuesday, bt);
            }

            if (checkBoxWed.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerWed1.Value.TimeOfDay,
                    StopTS = dateTimeWed2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Wednesday, bt);
            }

            if (checkBoxThurs.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerThurs1.Value.TimeOfDay,
                    StopTS = dateTimePickerThurs2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Thursday, bt);
            }

            if (checkBoxFri.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerFri1.Value.TimeOfDay,
                    StopTS = dateTimePickerFri2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Friday, bt);
            }

            if (checkBoxSat.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerSat1.Value.TimeOfDay,
                    StopTS = dateTimePickerSat2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Saturday, bt);
            }

            if (checkBoxSun.Checked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = dateTimePickerSun1.Value.TimeOfDay,
                    StopTS = dateTimePickerSun2.Value.TimeOfDay
                };
                AllowedTime.Add(DayOfWeek.Sunday, bt);
            }

            Globals.ThisAddIn._deferEmailsAllowedTime = AllowedTime;
            Globals.ThisAddIn.WriteVariables();

            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();            
        }


        private void CheckBox16_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn._deferEmailsAlwaysSendHighPriority = checkBox16.Checked;            
        }
    }
}
