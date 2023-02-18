using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Security.Principal;

[assembly : AssemblyTitle("PCONOFF")]
[assembly : AssemblyCompany("rostok - https://github.com/rostok/")]
[assembly : AssemblyVersion("1.0.2.0")]
[assembly : AssemblyFileVersion("1.0.2.0")]

namespace EventLogParser {
    class Event {
        public string Desc;
        public string Comment;
        public int Id;
        public DateTime Date;

        public Event(string onoff, string comment, int id, DateTime date) => (Desc, Comment, Id, Date)=(onoff, comment, id, date);
    }

    class Period {
        public DateTime Start;
        public DateTime Stop;
        public string Comment;

        public Period(DateTime start, DateTime stop, string comment="") => (Start, Stop, Comment)=(start, stop, comment);
        public Period() => (Start, Stop, Comment)=(DateTime.MinValue, DateTime.MinValue, "");
        public Period(string comment) => (Start, Stop, Comment)=(DateTime.MinValue, DateTime.MinValue, comment);
    }

    class Program {

        public static bool IsElevated()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static bool Elevate()
        {
            if (!IsElevated())
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.UseShellExecute = true;
                startInfo.WorkingDirectory = Environment.CurrentDirectory;
                startInfo.FileName = Process.GetCurrentProcess().MainModule.FileName;
                startInfo.Verb = "runas";
                startInfo.Arguments = Environment.CommandLine;

                try
                {
                    Process.Start(startInfo);
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }
            return true;
        }

        static List<Event> GetEvents(DateTime start, DateTime now) {
            List<Event> events = new List<Event>();
            List<EventLogQuery> queries = new List<EventLogQuery>();
            queries.Add(new EventLogQuery("System",   PathType.LogName, $"*[System[(Level=4 or Level=0) and TimeCreated[@SystemTime >= '{start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")}']]]"));
            queries.Add(new EventLogQuery("Security", PathType.LogName, $"*[System[(EventID>=4624 and EventID<=4801) and TimeCreated[@SystemTime >= '{start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")}']]]"));

            queries.ForEach(query=>{
                using(EventLogReader logReader = new EventLogReader(query)) {                                 
                    for (EventRecord eventRecord = logReader.ReadEvent(); eventRecord != null; eventRecord = logReader.ReadEvent()) {
                        if (eventRecord.TimeCreated == null) continue;
                        switch (eventRecord.Id) {
                            case 1:
                                events.Add(new Event("On", "Started/Woke up", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 6005:
                                events.Add(new Event("Off", "Stopped", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 42:
                                events.Add(new Event("Off", "Slept", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 13:
                                events.Add(new Event("Off", "Hibernated", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 12:
                                events.Add(new Event("On", "Resumed from hibernation", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4624:
                                events.Add(new Event("LogOn", "An account was successfully logged on", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4634:
                                events.Add(new Event("LogOff", "An account was logged off", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4647:
                                events.Add(new Event("LogOff", "User initiated logoff", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;                    
                            case 4778:
                                events.Add(new Event("LogOn", "A session was reconnected to a Window Station", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4779:
                                events.Add(new Event("LogOff", "A user disconnected a terminal server session without logging off", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4801:
                                events.Add(new Event("Unlocked", "The workstation was unlocked", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                            case 4800:
                                events.Add(new Event("Locked", "The workstation was locked", eventRecord.Id, eventRecord.TimeCreated.Value));
                                break;
                        }
                    }
                }
            });
            return events;
        } 

        public static List<Period> GetPeriods(List<Event> events, string start="On", string end="Off") {
            List<Period> periods = new List<Period>();
            Period currentPeriod = new Period(start);

            foreach (Event ev in events) {
                if (ev.Desc == start) {
                    if (currentPeriod.Start != DateTime.MinValue && currentPeriod.Stop != DateTime.MinValue) { 
                        periods.Add(currentPeriod);
                        currentPeriod = new Period(start);
                    }
                    if (currentPeriod.Start == DateTime.MinValue) currentPeriod.Start = ev.Date;
                }
                if (ev.Desc == end) {
                    if (currentPeriod.Start != DateTime.MinValue) {
                        currentPeriod.Stop = ev.Date;
                    }
                }
            }

            if (currentPeriod.Start != DateTime.MinValue) {
                if (currentPeriod.Stop == DateTime.MinValue) currentPeriod.Stop = DateTime.Now;
                periods.Add(currentPeriod);
            }

            periods = periods.SelectMany(p => p.Start.Date == p.Stop.Date ? new [] { p } : new [] {
                new Period(p.Start, new DateTime(p.Start.Year, p.Start.Month, p.Start.Day, 23, 59, 59), p.Comment), 
                new Period(new DateTime(p.Stop.Year, p.Stop.Month, p.Stop.Day, 0, 0, 0), p.Stop, p.Comment) }).ToList();

            return periods;
        }

        static void MakeHTML(List<Period> periods) {
            var periodsJS = string.Join(",\n", periods.Select(p => $"{{Start:new Date('{p.Start:yyyy-MM-ddTHH:mm:ss}'), Stop:new Date('{p.Stop:yyyy-MM-ddTHH:mm:ss}'), Comment:'{p.Comment}'}}"));
            File.WriteAllText("periods.js", $"var periods = [\n{periodsJS}\n];");

            var tableHTML = @"
<html>
    <head>
        <script src='periods.js'></script>
        <script>
            window.onload = function() {
                var table = document.getElementById('periods-table');
                var cd = '';
                periodCell = '';
                for (var i = 0; i < periods.length; i++) {
                    var start = periods[i].Start;
                    var stop = periods[i].Stop;
                    var comment = periods[i].Comment;
                    var startDay = start.toDateString();
                    var stopDay = stop.toDateString();
                    var id = start.toISOString().substring(0,10);
                    if (cd!=id) {
                    	cd = id;
                        var row = table.insertRow(-1);
                        var dateCell = row.insertCell(-1);
                        dateCell.innerHTML = id;
					    dateCell.style.width = '10vw';
                        periodCell = row.insertCell(-1);
					    periodCell.style.width = '90vw';
                    }
                    var width = (stop - start) / (24 * 60 * 60 * 1000);
                    var left = (start.getHours() * 60 + start.getMinutes()) / (24 * 60);
                    var outerdiv = document.createElement('div');
                    outerdiv.style.height = '100%';
                    outerdiv.style.width = '100%';
                    outerdiv.style.position = 'relative';

                    var div = document.createElement('div');
                    div.style.position = 'absolute';
                    div.style.opacity = '0.5';
                    div.style.backgroundColor = comment=='Unlocked'?'lightgreen':'lightblue';
                    div.style.height = '20px';
                    div.style.width = width * 100.1 + '%';
                    div.style.left = left * 100.1 + '%';
                    div.style.top = '-10px';
                    div.title = start.getHours()+':'+start.getMinutes()+' / '+stop.getHours()+':'+stop.getMinutes() +':'+comment;
                    outerdiv.appendChild(div);
                    periodCell.appendChild(outerdiv);
                }
            }
        </script>
    </head>
    <body>
        <table id='periods-table' border=1 style='border-collapse:collapse;'>
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Periods</th>
                </tr>
            </thead>
        </table>
    </body>
</html>
";
            File.WriteAllText("table.html", tableHTML);
        }

        public class PeriodDataGridView : DataGridView {
            private List<Period> periods;
            ToolTip tip = new ToolTip();

            public PeriodDataGridView(List<Period> periods) {
                this.periods = periods;
                InitializeComponent();
                this.DoubleBuffered = true;
                this.RowHeadersVisible = false;
                this.RowTemplate.DefaultCellStyle.BackColor = Color.White;
                this.RowTemplate.DefaultCellStyle.SelectionBackColor = Color.White;
                this.RowTemplate.DefaultCellStyle.SelectionForeColor = Color.Black;
            }

            private void InitializeComponent() {
                this.Columns.Add("Date", "Date");
                this.Columns[0].Width = this.Width / 10;
                this.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                this.Columns.Add("Periods", "Periods");
                this.Columns[1].Width = this.Width * 9 / 10;
                this.Columns[1].DefaultCellStyle.BackColor = Color.White;

                // this.Columns.Add("Hrs", "Hrs");
                // this.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                // this.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                this.RowTemplate.Height = 20;
                this.AllowUserToAddRows = false;

				var uniqueDates = new HashSet<string>(periods.Select(p=>p.Start.ToString("yyyy-MM-dd")));

                MouseMove += (object sender, MouseEventArgs e) => { Invalidate(); };

                foreach (var date in uniqueDates) {
                    var row = new DataGridViewRow();

                    var dateCell = new DataGridViewTextBoxCell();
                    dateCell.Value = date;

                    row.Cells.Add(dateCell);

                    // row.Cells.Add(new DataGridViewTextBoxCell());
                    // row.Cells.Add(new DataGridViewTextBoxCell());
                    // row.Cells[2].Value = "x";

                    dateCell = new DataGridViewTextBoxCell();
                    this.Rows.Add(row);
                }
            }

            protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e) {
                base.OnCellPainting(e);

                if (e.RowIndex < 0 || e.ColumnIndex != 1) return;

                e.Graphics.FillRectangle(Brushes.White, e.CellBounds);

                var row = this.Rows[e.RowIndex];
                var date = (string) row.Cells[0].Value;
				Brush brush;

                var rects = new List<Rectangle>();
                Point mousePos = this.PointToClient(Control.MousePosition);
                Rectangle selectedRect = new Rectangle();
                string selectedText = "";
                foreach (var period in periods) {
                    if (period.Start.ToString("yyyy-MM-dd") != date) continue;

                    var width = (period.Stop - period.Start).TotalMilliseconds / (24 * 60 * 60 * 1000);
                    var left = (period.Start.Hour * 60 + period.Start.Minute) / (24 * 60.0f);

                    var rect = new Rectangle(
                        (int) (e.CellBounds.Left + e.CellBounds.Width * left),
                        e.CellBounds.Top,
                        (int) (e.CellBounds.Width * width),
                        e.CellBounds.Height - 1);

                    rects.Add(rect);
                    brush = new SolidBrush(Color.FromArgb(128, Color.LightBlue));
                    if (period.Comment=="Unlocked") brush = new SolidBrush(Color.FromArgb(128, Color.LightGreen));
                    e.Graphics.FillRectangle(brush, rect);
                    if (rect.Contains(mousePos)) {
                        brush = new SolidBrush(Color.FromArgb(128, Color.DeepSkyBlue));
                        if (period.Comment=="Unlocked") brush = new SolidBrush(Color.FromArgb(128, Color.Green));
                        e.Graphics.FillRectangle(brush, rect);
                        selectedRect = rect;
                        selectedText = $"{period.Start.Hour}:{period.Start.Minute} / {period.Stop.Hour}:{period.Stop.Minute} : {period.Comment}";
                    }
                }

                //e.Graphics.DrawRectangles(new Pen(Brushes.Black), rects.ToArray());

                
                if (selectedRect.Width > 0) {
                    Font font = new Font("Arial Narrow", 10);
                    brush = Brushes.Black;
                    e.Graphics.DrawString(selectedText, font, brush, new PointF(selectedRect.Left, selectedRect.Top));
                    //tip.Show(selectedText, this.Parent, mousePos.X+32, mousePos.Y+40);
                };

                e.Handled = true;
            }
        }

        static void CreateDataGridView(List<Period> periods) {
            Form form = new Form();
            form.Text = "PCOnOff";
            form.WindowState = FormWindowState.Maximized;
            form.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            PeriodDataGridView dgv = new PeriodDataGridView(periods);
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Dock = DockStyle.Fill;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            form.Resize += (sender, args) => {
                dgv.Columns[0].Width = form.Width / 10;
                dgv.Columns[1].Width = form.Width * 9 / 10;
            };

            if (dgv.RowCount>0) dgv.CurrentCell = dgv[1, dgv.RowCount - 1];

            EventHandler Init = new EventHandler((sender, args) => { });

            form.Activated += Init;
            form.Shown += Init;

            form.Controls.Add(dgv);
            form.ShowDialog();
        }

        static void Main(string[] args) {
            if (args.Contains("-h") || args.Contains("/?") || args.Contains("--help")) {
                Console.WriteLine("PCONOFF will show when this computer was turned on and unlocked.");
                Console.WriteLine("syntax: pconoff [days] [-l] [-k] [-c] [-t]");
                Console.WriteLine(" days    how deep in time should the events be analyzed (90 by default)");
                Console.WriteLine(" -l      will investigate log on/off events");
                Console.WriteLine(" -k      will investigate unlocked/locked events");
                Console.WriteLine(" -c      will print report to console");
                Console.WriteLine(" -v      verbose output of events");
                Console.WriteLine(" -t      will generate periods.js and table.html & will NOT show gird");
                Console.WriteLine("");
                Console.WriteLine("this comes with MIT license from rostok - https://github.com/rostok/");
                return;
            }

            if (!IsElevated()) {
                Elevate();
                return;
            }

            int daysBack = 90;
            if (args.Length>0 && !Int32.TryParse(args[0], out daysBack)) daysBack = 90;
            DateTime now = DateTime.Now;
            DateTime start = now.AddDays(-daysBack);
            var events = GetEvents(start, now);

            if (args.Contains("-v"))
                events.ForEach(e=>Console.WriteLine($"{e.Date} {e.Id} {e.Desc} {e.Comment}"));

            List<Period> periods;
            if (args.Contains("-l")) 
            	periods = GetPeriods(events, "On", "Off");
            else
                if (args.Contains("-k")) 
            	    periods = GetPeriods(events, "Unlocked", "Locked");
            	else
                    periods = GetPeriods(events, "On", "Off").Concat(GetPeriods(events, "Unlocked", "Locked")).OrderBy(p=>p.Start).ToList();

            if (args.Contains("-c")) 
                periods.ForEach(p=>Console.WriteLine($"Start:{p.Start.ToString("yyyy-MM-dd HH:mm:ss")}, Stop:{p.Stop.ToString("yyyy-MM-dd HH:mm:ss")}, Comment:{p.Comment}"));

            if (args.Contains("-t"))
                MakeHTML(periods);

            if (!args.Contains("-t")&&!args.Contains("-c"))
                CreateDataGridView(periods);
        }
    }
}