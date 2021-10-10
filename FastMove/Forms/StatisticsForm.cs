using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Globalization;



namespace FastMove
{
    public partial class StatisticsForm : Form
    {

        public string Pad(int i)
        {
            if (i < 10)
            {
                return "0" + i;
            }
            return i + "";
        }

        public void DrawChart1()
        {
            /* 
          * Chart 1
          */
            chart1.Series.Clear();
            chart1.ChartAreas.Clear();

            var chartArea = new ChartArea();
            chartArea.AxisX.LabelStyle.Format = "dd/MMM\nhh:mm";
            chartArea.AxisX.MajorGrid.LineColor = Color.LightGray;
            chartArea.AxisY.MajorGrid.LineColor = Color.LightGray;
            chartArea.AxisX.LabelStyle.Font = new Font("Consolas", 8);
            chartArea.AxisY.LabelStyle.Font = new Font("Consolas", 8);
            chart1.ChartAreas.Add(chartArea);

            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;

            // Set palette.
            this.chart1.Palette = ChartColorPalette.SeaGreen;

            // Set title.
            this.chart1.Titles.Add("Mail(s) per day");

            Series series = this.chart1.Series.Add("Date");
            series.ChartType = SeriesChartType.Line;
            series.XValueType = ChartValueType.DateTime;

            var list = _MailsPerDay.Keys.ToList();
            list.Sort();

            List<int> Values = new List<int>();

            foreach (var key in list)
            {
                Values.Add(_MailsPerDay[key]);
            }

            chart1.Series["Date"].Points.DataBindXY(list, Values);

            // draw!
            chart1.Invalidate();
        }

        public void DrawChart2()
        {

            /*
             * Chart 2 - Mails per Week
             * 
             */
            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;
            Dictionary<string, int> _MailsPerWeek = new Dictionary<string, int>();
            var chartArea2 = new ChartArea();

            chart2.Series.Clear();
            chart2.ChartAreas.Clear();

            chartArea2.AxisX.LabelStyle.Format = "dd/MMM\nhh:mm";
            chartArea2.AxisX.MajorGrid.LineColor = Color.LightGray;
            chartArea2.AxisY.MajorGrid.LineColor = Color.LightGray;
            chartArea2.AxisX.LabelStyle.Font = new Font("Consolas", 8);
            chartArea2.AxisY.LabelStyle.Font = new Font("Consolas", 8);
            chart2.ChartAreas.Add(chartArea2);

            // Set palette.
            this.chart2.Palette = ChartColorPalette.SeaGreen;

            // Set title.
            this.chart2.Titles.Add("Mail(s) per week");

            Series series2 = this.chart2.Series.Add("Week");
            series2.ChartType = SeriesChartType.Line;
            series2.XValueType = ChartValueType.DateTime;

            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            string weekNr;

            List<DateTime> list2 = _MailsPerDay.Keys.ToList();
            list2.Sort();

            foreach (var key in list2)
            {
                weekNr = cal.GetYear(key).ToString() + "-" +
                    cal.GetWeekOfYear(key, dfi.CalendarWeekRule, dfi.FirstDayOfWeek).ToString();
                if (_MailsPerWeek.ContainsKey(weekNr))
                {
                    _MailsPerWeek[weekNr] += _MailsPerDay[key];
                }
                else
                {
                    _MailsPerWeek.Add(weekNr, _MailsPerDay[key]);
                }
            }


            //-
            var list3 = _MailsPerWeek.Keys.ToList();
            list3.Sort();

            List<int> Values2 = new List<int>();

            foreach (var key in list3)
            {
                Values2.Add(_MailsPerWeek[key]);
            }

            chart2.Series["Week"].Points.DataBindXY(list3, Values2);

            // draw!
            chart2.Invalidate();

        }

        public StatisticsForm()
        {
            InitializeComponent();

            double seconds = Globals.ThisAddIn._InboxAvg;
            TimeSpan TS = TimeSpan.FromSeconds(seconds);
            string AvgText = TS.Days + " days," +
                Pad(TS.Hours) + " hours, " +
                Pad(TS.Minutes) + " minutes " +
                Pad(TS.Seconds) + " seconds";

            textBox1.Text = AvgText;

            seconds = 0;
            int count = 0;
            foreach (double d in Globals.ThisAddIn._avgTimeBeforeMove)
            {
                seconds += d;
                count++;
            }
            if (count > 0)
                seconds /= count;

            TS = TimeSpan.FromSeconds(seconds);
            AvgText = TS.Days + " days," +
                Pad(TS.Hours) + " hours, " +
                Pad(TS.Minutes) + " minutes " +
                Pad(TS.Seconds) + " seconds";

            textBox2.Text = AvgText;

            DrawChart1();
            DrawChart2();

         
            /*
             * Update datagrid
             */
            Dictionary<DateTime, int> _MailsPerDay = Globals.ThisAddIn._MailsPerDay;

            DataTable table = new DataTable();
            table.Columns.Add("From", typeof(string));
            table.Columns.Add("Count", typeof(int));

            foreach (var pair in Globals.ThisAddIn._MailsFromWho)
            {                
                table.Rows.Add(pair.Key,pair.Value);
            }

            dataGridView1.DataSource = table;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.ClearSelection();

            DataGridViewColumn column = dataGridView1.Columns[0];
            column.Width = 400;            
            DataGridViewColumn column2 = dataGridView1.Columns[1];
            column2.Width = 100;

            this.dataGridView1.Sort(this.dataGridView1.Columns[1], ListSortDirection.Descending);


            // Update statusbar!
            DateTime day = DateTime.Today.Date;
            count = 0;

            if (_MailsPerDay.ContainsKey(day))
            {
                count = _MailsPerDay[day];
            } 
            toolStripStatusLabel1.Text = string.Format("Today: {0}", count);

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

            
            toolStripStatusLabel2.Text = string.Format("Last week: {0}", count);
            statusStrip1.Refresh();
            toolStripStatusLabel4.Text = string.Format("Version: {0}", Globals.ThisAddIn.publishedVersion);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
            var form1 = (Form1)Tag;
            form1.Show();            
        }

        private void StatisticsForm_Load(object sender, EventArgs e)
        {

        }    


    }
}
