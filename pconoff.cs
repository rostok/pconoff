using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace EventLogParser {
    class Event {
        public string OnOff { get; set; }
        public string Comment { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }

        public Event(string onoff, string comment, int id, DateTime date) {
            OnOff = onoff;
            Comment = comment;
            Id = id;
            Date = date;
        }
    }

    class Program {

        public static List < (DateTime start, DateTime stop) > GetPeriods(List<Event> events) {
            List < (DateTime start, DateTime stop) > periods = new List < (DateTime start, DateTime stop) > ();
            (DateTime start, DateTime stop) currentPeriod = (DateTime.MinValue, DateTime.MinValue);

            foreach (Event ev in events) {
                if (ev.OnOff == "On") {
                    if (currentPeriod.start == DateTime.MinValue) {
                        currentPeriod.start = ev.Date;
                    } else {
                        //periods.Add(currentPeriod);
                        //currentPeriod = (ev.Date, DateTime.MinValue);
                    }
                } else {
                    if (currentPeriod.start != DateTime.MinValue) {
                        currentPeriod.stop = ev.Date;
                        periods.Add(currentPeriod);
                        currentPeriod = (DateTime.MinValue, DateTime.MinValue);
                    }
                }
            }

            if (currentPeriod.start != DateTime.MinValue && currentPeriod.stop != DateTime.MinValue) {
                periods.Add(currentPeriod);
            }

            periods = periods.SelectMany(p => p.start.Date == p.stop.Date ? new [] { p } : new [] {
                (p.start, new DateTime(p.start.Year, p.start.Month, p.start.Day, 23, 59, 59)), (new DateTime(p.stop.Year, p.stop.Month, p.stop.Day, 0, 0, 0), p.stop) }).ToList();

            return periods;
        }

        static void MakeHTML(List < (DateTime start, DateTime stop) > periods) {
            var periodsJS = string.Join(",\n", periods.Select(p => $"{{start: new Date('{p.start:yyyy-MM-ddTHH:mm:ss}'), stop: new Date('{p.stop:yyyy-MM-ddTHH:mm:ss}')}}"));
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
                    var start = periods[i].start;
                    var stop = periods[i].stop;
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
                    div.style.backgroundColor = 'lightblue';
                    div.style.height = '20px';
                    div.style.width = width * 100.1 + '%';
                    div.style.left = left * 100.1 + '%';
                    div.style.top = '-10px';
                    div.title = start.getHours()+':'+start.getMinutes()+' / '+stop.getHours()+':'+stop.getMinutes();
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
            private List < (DateTime start, DateTime stop) > periods;

            public PeriodDataGridView(List < (DateTime start, DateTime stop) > periods) {
                this.periods = periods;
                InitializeComponent();
                this.DoubleBuffered = true;
                this.RowTemplate.DefaultCellStyle.BackColor = Color.White;
                this.RowTemplate.DefaultCellStyle.SelectionBackColor = Color.White;
                this.RowTemplate.DefaultCellStyle.SelectionForeColor = Color.Black;
            }

            private void InitializeComponent() {
                this.Columns.Add("Date", "Date");
                this.Columns[0].Width = this.Width / 10;
                this.Columns.Add("Periods", "Periods");
                this.Columns[1].Width = this.Width * 9 / 10;
                this.Columns[1].DefaultCellStyle.BackColor = Color.White;
                this.RowTemplate.Height = 20;

                var uniqueDates = new HashSet<string>();
                foreach (var period in periods) {
                    uniqueDates.Add(period.start.ToString("yyyy-MM-dd"));
                }

                MouseMove += (object sender, MouseEventArgs e) => { Invalidate(); };

                foreach (var date in uniqueDates) {
                    var row = new DataGridViewRow();

                    var dateCell = new DataGridViewTextBoxCell();
                    dateCell.Value = date;

                    row.Cells.Add(dateCell);
                    dateCell = new DataGridViewTextBoxCell();
                    //			dateCell.Value = "x";
                    //            row.Cells.Add(dateCell);

                    //            var periodCell = new DataGridViewCell();
                    //            row.Cells[1] = periodCell;
                    this.Rows.Add(row);
                }
            }

            protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e) {
                base.OnCellPainting(e);

                if (e.RowIndex < 0 || e.ColumnIndex != 1) {
                    return;
                }

                e.Graphics.FillRectangle(Brushes.White, e.CellBounds);

                var row = this.Rows[e.RowIndex];
                var date = (string) row.Cells[0].Value;

                //row.Cells[1].Value = "";
                var rects = new List<Rectangle>();
                Point mousePos = this.PointToClient(Control.MousePosition);
                Rectangle selectedRect = new Rectangle();
                string selectedText = "";
                foreach (var period in periods) {
                    if (period.start.ToString("yyyy-MM-dd") != date) continue;

                    var width = (period.stop - period.start).TotalMilliseconds / (24 * 60 * 60 * 1000);
                    var left = (period.start.Hour * 60 + period.start.Minute) / (24 * 60.0f);
                    //            var left = (period.start - period.start.Date).TotalMilliseconds / (24 * 60 * 60 * 1000); 

                    var rect = new Rectangle(
                        (int) (e.CellBounds.Left + e.CellBounds.Width * left),
                        e.CellBounds.Top,
                        (int) (e.CellBounds.Width * width),
                        e.CellBounds.Height - 1);

                    rects.Add(rect);
                    e.Graphics.FillRectangle(Brushes.LightBlue, rect);
                    //            e.Graphics.DrawRectangle(new Pen(Brushes.Black), rect);
                    if (rect.Contains(mousePos)) {
                        e.Graphics.FillRectangle(Brushes.DeepSkyBlue, rect);
                        selectedRect = rect;
                        selectedText = $"{period.start.Hour}:{period.start.Minute} / {period.stop.Hour}:{period.stop.Minute}";
                    }

                    //row.Cells[1].Value += period.start.ToString();
                }
                //        e.Graphics.FillRectangles(Brushes.LightBlue, rects.ToArray());
                e.Graphics.DrawRectangles(new Pen(Brushes.Black), rects.ToArray());

                if (selectedRect.Width > 0) {
                    Font font = new Font("Arial", 12);
                    Brush brush = Brushes.Black;
                    e.Graphics.DrawString(selectedText, font, brush, new PointF(selectedRect.Left, selectedRect.Top));
                }

                e.Handled = true;
            }
        }

        static void CreateDataGridView(List < (DateTime start, DateTime stop) > periods) {
            Form form = new Form();
            form.Text = "PCOnOff";
            form.WindowState = FormWindowState.Maximized;

            PeriodDataGridView dgv = new PeriodDataGridView(periods);
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Dock = DockStyle.Fill;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            form.Resize += (sender, args) => {
                dgv.Columns[0].Width = form.Width / 10;
                dgv.Columns[1].Width = form.Width * 9 / 10;
            };

            dgv.CurrentCell = dgv[1, dgv.RowCount - 1];

            EventHandler Init = new EventHandler((sender, args) => { });

            form.Activated += Init;
            form.Shown += Init;

            form.Controls.Add(dgv);
            form.ShowDialog();
        }

        static List<Event> GetEvents(DateTime start, DateTime now) {
            List<Event> events = new List<Event>();
            EventLogQuery query = new EventLogQuery("System", PathType.LogName, "*[System[(Level=4 or Level=0) and TimeCreated[@SystemTime >= '" + start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ") + "']]]");
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
                    }
                }
            }
            return events;
        } 

        static void Main(string[] args) {
            if (args.Contains("-h") || args.Contains("/?") || args.Contains("--help")) {
                Console.WriteLine("pconoff will show when this computer was turned on.");
                Console.WriteLine("syntax: pconoff [days-back] [-t] [-h]");
                Console.WriteLine(" days    how deep in time should the events be analyzed (90 by default)");
                Console.WriteLine(" -c      will print report to console");
                Console.WriteLine(" -t      will generate periods.js and table.html & will NOT show gird");
                Console.WriteLine("");
                Console.WriteLine("this comes with MIT license from rostok - https://github.com/rostok/");
                return;
            }

            int daysBack = 90;
            if (args.Length>0 && !Int32.TryParse(args[0], out daysBack)) daysBack = 90;
            DateTime now = DateTime.Now;
            DateTime start = now.AddDays(-daysBack);
            var events = GetEvents(start, now);

            var periods = GetPeriods(events);

            if (args.Contains("-c"))
                foreach ((DateTime start, DateTime stop) period in periods) {
                    Console.WriteLine("Start: " + period.start.ToString("yyyy-MM-dd HH:mm:ss") + ", Stop: " + period.stop.ToString("yyyy-MM-dd HH:mm:ss"));
                }

            if (args.Contains("-t"))
                MakeHTML(periods);

            if (!args.Contains("-t"))
                CreateDataGridView(periods);
        }
    }
}