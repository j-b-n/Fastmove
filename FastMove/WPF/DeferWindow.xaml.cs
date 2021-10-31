using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FastMove.WPF
{

    public class ClockValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {            
            string str = value as string;

            try
            {
                TimeSpan ts = TimeSpan.Parse(str);                
            }
            catch (FormatException)
            {
                return new ValidationResult(false, str+": Bad Format");
            }
            catch (OverflowException)
            {
                return new ValidationResult(false, str + ": Overflow");                
            }                                               

            return new ValidationResult(true, null);
        }
    }

    /// <summary>
    /// Interaction logic for DeferWindow.xaml      
    /// </summary>
    public partial class DeferWindow : AdonisUI.Controls.AdonisWindow
    {
        ThemeManager themeManager = null;

        public String _MondayFrom { get; set; }
        public String _MondayTo { get; set; }
        public String _TuesdayFrom { get; set; }
        public String _TuesdayTo { get; set; }
        public String _WednesdayFrom { get; set; }
        public String _WednesdayTo { get; set; }
        public String _ThursdayFrom { get; set; }
        public String _ThursdayTo { get; set; }
        public String _FridayFrom { get; set; }
        public String _FridayTo { get; set; }
        public String _SaturdayFrom { get; set; }
        public String _SaturdayTo { get; set; }
        public String _SundayFrom { get; set; }
        public String _SundayTo { get; set; }
        
        public DeferWindow()
        {
            InitializeComponent();
            DataContext = this;

            AdonisUI.SpaceExtension.SetSpaceResourceOwnerFallback(this);
            themeManager = new ThemeManager(this);
            themeManager.SetDefaultTheme();            
        }
        private void ToogleThemeBtn_Click(object sender, RoutedEventArgs e)
        {
            themeManager.ToggleTheme();

            if (themeManager.GetTheme() == ResourceLocator.DarkColorScheme)
                ToogleThemeBtn.Content = "☀️";
            else
                ToogleThemeBtn.Content = "🌙";

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            DeferEmailsCB.IsChecked = Globals.ThisAddIn._deferEmails;
            AlwaysSend.IsChecked = Globals.ThisAddIn._deferEmailsAlwaysSendHighPriority;

            if (Globals.ThisAddIn._deferEmails)
            {
                DateTime deferTime = Globals.ThisAddIn.NextPossibleSendTime();
                NextPossibleSendTime.Text = string.Format("Next possible send time: {0}", deferTime.ToString());
            }
            else
            {
                NextPossibleSendTime.Text = string.Format("Next possible send time: Immediately");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Monday))
            {
                MondayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Monday];
                MondayFrom.Text =  BT.StartTS.ToString(@"hh\:mm");
                MondayTo.Text= BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Tuesday))
            {
                TuesdayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Tuesday];
                TuesdayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                TuesdayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Wednesday))
            {
                WednesdayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Wednesday];
                WednesdayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                WednesdayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Thursday))
            {
                ThursdayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Thursday];
                ThursdayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                ThursdayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Friday))
            {
                FridayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Friday];
                FridayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                FridayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Saturday))
            {
                SaturdayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Saturday];
                SaturdayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                SaturdayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }

            if (Globals.ThisAddIn._deferEmailsAllowedTime.ContainsKey(DayOfWeek.Sunday))
            {
                SundayCB.IsChecked = true;
                BetweenTime BT = Globals.ThisAddIn._deferEmailsAllowedTime[DayOfWeek.Sunday];
                SundayFrom.Text = BT.StartTS.ToString(@"hh\:mm");
                SundayTo.Text = BT.StopTS.ToString(@"hh\:mm");
            }
            ///Update StatusBar content

            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            DateTime day = DateTime.Today.Date;
            int count = 0;

            if (_MailsPerDay.ContainsKey(day))
            {
                count = _MailsPerDay[day];
            }
            SBToday.Text = string.Format("Today: {0}", count);

            //Last week
            count = 0;

            for (int i = 0; i > -6; i--)
            {
                if (_MailsPerDay.ContainsKey(day))
                {
                    count += _MailsPerDay[day];
                }
                day = day.AddDays(i);
            }

            SBLastWeek.Text = string.Format("Last week: {0}", count);
            SBVersion.Text = string.Format("Version: {0}", Globals.ThisAddIn.publishedVersion);
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            //Dictionary Weekday with list of BetweenTime
            //BetweenTime struct timespanStart and timespanStop                        

            Dictionary<DayOfWeek, BetweenTime> AllowedTime = new Dictionary<DayOfWeek, BetweenTime>();

            if (MondayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_MondayFrom),
                    StopTS = ConvertToTimeSpan(_MondayTo)
                };
                AllowedTime.Add(DayOfWeek.Monday, bt);
            }

            if (TuesdayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_TuesdayFrom),
                    StopTS = ConvertToTimeSpan(_TuesdayTo)
                };
                AllowedTime.Add(DayOfWeek.Tuesday, bt);
            }

            if (WednesdayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_WednesdayFrom),
                    StopTS = ConvertToTimeSpan(_WednesdayTo)
                };
                AllowedTime.Add(DayOfWeek.Wednesday, bt);
            }

            if (ThursdayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_ThursdayFrom),
                    StopTS = ConvertToTimeSpan(_ThursdayTo)
                };
                AllowedTime.Add(DayOfWeek.Thursday, bt);
            }

            if (FridayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_FridayFrom),
                    StopTS = ConvertToTimeSpan(_FridayTo)
                };
                AllowedTime.Add(DayOfWeek.Friday, bt);
            }
            if (SaturdayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_SaturdayFrom),
                    StopTS = ConvertToTimeSpan(_SaturdayTo)
                };
                AllowedTime.Add(DayOfWeek.Saturday, bt);
            }
            if (SundayCB.IsChecked == true)
            {
                BetweenTime bt = new BetweenTime
                {
                    StartTS = ConvertToTimeSpan(_SundayFrom),
                    StopTS = ConvertToTimeSpan(_SundayTo)
                };
                AllowedTime.Add(DayOfWeek.Sunday, bt);
            }

            Globals.ThisAddIn._deferEmailsAllowedTime = AllowedTime;
            Globals.ThisAddIn.WriteVariables();
            this.Close();
        }

        private void DeferEmailsCB_Checked(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn._deferEmails = (bool)DeferEmailsCB.IsChecked;

            if (Globals.ThisAddIn._deferEmails)
            {
                DateTime deferTime = Globals.ThisAddIn.NextPossibleSendTime();
                NextPossibleSendTime.Text = string.Format("Next possible send time: {0}", deferTime.ToString());
            }
            else
            {
                NextPossibleSendTime.Text = string.Format("Next possible send time: Immediately");
            }            
        }

        public TimeSpan ConvertToTimeSpan(String str)
        {            
            try
            {
                TimeSpan ts = TimeSpan.Parse(str);
                return ts;
            }            
            catch 
            {
                return new TimeSpan();
            }            
        }        
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }    
}
