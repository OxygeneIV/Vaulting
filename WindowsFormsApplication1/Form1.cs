using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Color = System.Drawing.Color;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Threading;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        private static string root;
        private static string resultfile;
        private static string horseresultfile;

        private static string startlistfile;
        private static string inboxFolder;
        private static string outboxFolder;
        private static string fakeboxFolder;
        private static string backupFolder;
        private static bool showmessageboxes;
        private static bool competitionstarted;


        public static string sortedresultsfile;
        public static string omvandfile;
        public static string printedresultsFolder;
        public static string mergedresultsFolder;
        public static string horseResultsFolder;
        public static string htmlResultsFolder;
        public static string cssFolder;
        public static string logosFolder;
        private static string workingDirectory;
        private static bool fake;
        private static string fakefile;

        static object lockObject = new object();

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            setPathes();
         
            dataGridView1.AutoGenerateColumns = true;
            dataGridView2.AutoGenerateColumns = true;
            dataGridView3.AutoGenerateColumns = true;
            dataGridView3.RowPrePaint += new DataGridViewRowPrePaintEventHandler(dataGridView3_RowPrePaint);
            tabPage1.Text = "Klasser";
            tabPage2.Text = "Deltagare";
            tabPage3.Text = "Resultat";
        }

        private void setPathes()
        {

            try
            {
                root = ConfigurationManager.AppSettings["root"];
                fake = bool.Parse(ConfigurationManager.AppSettings["fake"]);
                showmessageboxes = bool.Parse(ConfigurationManager.AppSettings["showmessageboxes"]);
                competitionstarted = bool.Parse(ConfigurationManager.AppSettings["competitionstarted"]);
              
                workingDirectory = string.IsNullOrEmpty(root) ? Application.StartupPath : root;

                if (!Directory.Exists(workingDirectory))
                {
                    throw new Exception("Failed to find working directory " + workingDirectory + "\n" + " App.Config 'root' set to " + root);
                }

              this.Text = $"Voltigeresultat - {workingDirectory}";

              if (!fake)
                    buttonFakeResults.Enabled = false;

                if(competitionstarted)
                {
                    buttonPopulateSheetsWithVaulters.Enabled = false;
                    buttonCreateResultSheets.Enabled = false;
                }

                // Folders
                List<string> foldersToCreate = new List<string>();

                inboxFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["inbox"]);
                foldersToCreate.Add(inboxFolder);

                fakeboxFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["fakebox"]);
                foldersToCreate.Add(fakeboxFolder);

                backupFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["backup"]);
                foldersToCreate.Add(backupFolder);

                outboxFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["outbox"]);
                foldersToCreate.Add(outboxFolder);

                logosFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logos"]);

                if(!Directory.Exists(logosFolder))
                    logosFolder = Path.Combine(Application.StartupPath, "logos");

                //foldersToCreate.Add(logosfolder);

                printedresultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["printedresults"]);
                foldersToCreate.Add(printedresultsFolder);

                htmlResultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["htmlresultsfolder"]);
                foldersToCreate.Add(htmlResultsFolder);

                cssFolder = Path.Combine(htmlResultsFolder, ConfigurationManager.AppSettings["cssfolder"]);
                foldersToCreate.Add(cssFolder);

          

                mergedresultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["mergedresults"]);
                 foldersToCreate.Add(mergedresultsFolder);

                horseResultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["horseresultsfolder"]);
                foldersToCreate.Add(horseResultsFolder);
       

              foreach (var folder in foldersToCreate)
                {
                    var dirinfo = Directory.CreateDirectory(folder);
     
                }

                // Files
                resultfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["results"]);
                horseresultfile = Path.Combine(horseResultsFolder, ConfigurationManager.AppSettings["horseresults"]);
                startlistfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["startlist"]);
                sortedresultsfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["sortedresults"]);
                omvandfile= Path.Combine(workingDirectory, ConfigurationManager.AppSettings["omvandstartordning"]);
                //ridsportlogo = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logo"]);
                //preliminaryResults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["prel"]);
                //logovoid = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logovoid"]);

                String cssfile = Path.Combine(Environment.CurrentDirectory, "html/stylesheet.css");
                File.Copy(cssfile, Path.Combine(cssFolder, "stylesm.css"), true);

                var files= Directory.EnumerateFiles(Path.Combine(Environment.CurrentDirectory, "html/img"));
                foreach(String f in files)
                {
                   
                    File.Copy(f, htmlResultsFolder+"/" +Path.GetFileName(f), true);
                }



                fakefile = Path.Combine(fakeboxFolder, "fakedresults.xlsx");

                if(!File.Exists(resultfile))
                {
                    showMessageBox("First time using folder " + workingDirectory + ". Copying base result file");
                    var ff = Path.Combine(Application.StartupPath, ConfigurationManager.AppSettings["results"]);
                    File.Copy(ff, resultfile);
                }

                //if (!File.Exists(ridsportlogo))
                //{
                //    var logoFile = Path.Combine(Application.StartupPath,"logos", ConfigurationManager.AppSettings["logo"]);
                //    File.Copy(logoFile, ridsportlogo);
                //}
                //if (!File.Exists(logovoid))
                //{
                //    var logoFile = Path.Combine(Application.StartupPath, "logos", ConfigurationManager.AppSettings["logovoid"]);
                //    File.Copy(logoFile, logovoid);
                //}

                //if (!File.Exists(preliminaryResults))
                //{
                //    var preliminaryResultsFile = Path.Combine(Application.StartupPath, "logos", ConfigurationManager.AppSettings["prel"]);
                //    File.Copy(preliminaryResultsFile, preliminaryResults);
                //}

            }
            catch (Exception e)
            {
                UpdateMessageTextBox(e.Message);
                Application.Exit();
            }
        }


        // read classes from startlist
        private List<Klass> readClasses()
        {
            if(!File.Exists(startlistfile))
            {
              UpdateMessageTextBox($"No startlist found, expecting " + startlistfile);
              return new List<Klass>();
            }

            // Klasser
            FileInfo startlist = new FileInfo(startlistfile);
            List<Klass> classes;

            using (var pck = new ExcelPackage(startlist))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets["Klasser"];
                List<ExcelRange> rows = new List<ExcelRange>();
                List<ExcelRow> truerows = new List<ExcelRow>();
                List<List<string>> cellvals = new List<List<string>>();
                Dictionary<int, List<string>> cellvalues = new Dictionary<int, List<string>>();
                var dim = ws.Dimension;
                
                for (var rowNum = 1; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    object[,]  dict;   
                    var trueRow = ws.Row(rowNum);
                    
                    var row = ws.Cells[string.Format("{0}:{0}", rowNum)];

                    dict = (object[,])row.Value;
                    var tmplist = new List<string>();
                    for (int i = 0; i < ws.Dimension.End.Column; i++)
                    {

                        tmplist.Add(dict[0, i] == null ? "" : (Convert.ToString(dict[0, i])).Trim());
                    }

                    string text = row.ElementAt(0).Text;
                    bool allEmpty = text == "";
                    float newSmclass;
                    bool isfloat = float.TryParse(text,NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out newSmclass);
                    if (allEmpty || !(text.IsNumeric() || isfloat)) continue; // skip this row
                    rows.Add(row);
                    truerows.Add(trueRow);
                    cellvals.Add(tmplist);
                }                
                classes = cellvals.Select(r => Klass.RowToClass(r)).ToList();
         
                // Remove .2-classes SM & NM
                classes.RemoveAll(c => c.Name.EndsWith(".2"));
            }
          UpdateMessageTextBox($"Found {classes.Count} classes");
          return classes;
        }

        // read deltagare from startlist
        private List<Deltagare> readVaulters()
        {
            if (!File.Exists(startlistfile))
            {
                UpdateMessageTextBox($"No startlist found, expecting " + startlistfile);
                return new List<Deltagare>();
            }

            FileInfo startlist = new FileInfo(startlistfile);
            List<Deltagare> deltagare;

            using (var pck = new ExcelPackage(startlist))
            {
                ExcelWorksheet starters = pck.Workbook.Worksheets["Deltagare"];

                List<ExcelRange> deltagarlista = new List<ExcelRange>();

                for (var rowNum = 1; rowNum <= starters.Dimension.End.Row; rowNum++)
                {
                    var row = starters.Cells[string.Format("{0}:{0}", rowNum)];
                    string text = row.ElementAt(0).Text;
                    bool allEmpty = text == "";
                    float newSmclass;
                    bool isfloat = float.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out newSmclass);

                    if (allEmpty || !(text.IsNumeric() || isfloat) ) continue; // skip this row
                    deltagarlista.Add(row);
                }
                deltagare = deltagarlista.Select(r => Deltagare.RowToClass(r)).ToList();


            }

            List<Deltagare> deltagare2 = new List<Deltagare>();

            // Double .2-people SM & NM

            foreach (var d in deltagare)
            {
                

                if (d.Klass.EndsWith(".2"))
                {
                    var d1 = d.Duplicate();
                    d1.Klass = d.Klass.Replace(".2", "");
                    d1.Id = d.Id.Replace(".2", "");
                    deltagare2.Add(d1);

                    var d2 = d.Duplicate();
                    d2.Klass = d2.Klass.Replace(".2", ".1");
                    d2.Id = d2.Id.Replace(".2", ".1");
                    deltagare2.Add(d2);
                }
                else
                {
                    deltagare2.Add(d.Duplicate());
                }
            }

          var distinctIds = deltagare2.Select(d => d.Id).Distinct().Count();
          var duplicates = deltagare2.Count - distinctIds;

             UpdateMessageTextBox($"Found {deltagare2.Count} vaulters, {duplicates} duplicate IDs" );
             return deltagare2;
        }


        /// <summary>
        /// convert worksheet to DataTable
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        private DataTable ToDataTable(ExcelWorksheet ws, bool hasHeaderRow = false)
        {
            var tbl = new DataTable();
            for (int i = 0; i < 15; i++)
            {
                tbl.Columns.Add();
            }

            var startRow = 1;

            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, 15];
                var row = tbl.NewRow();
                var items = wsRow.ToArray().Select(p => p.Text);
                row.ItemArray = items.ToArray();
                tbl.Rows.Add(row);
            }
            return tbl;
        }


        class ClassSelect
        {
            public string Text { get; set; }
            public string Value { get; set; }
        }

        // populate Tabs with vaulters and classes
        private void ReadClassesAndVaultersFromStartlist()
        {

            List<Klass> classes = readClasses();
            var ccount = classes.Count;
            tabControl1.TabPages[1].Text = " Classes (" + ccount + ")";

            var classesInGrid = classes
                .Select(r => new
                {
                    Name = r.Name,
                    Team = r.Description,
                    Moment1 = r.Moments.Count >= 1 ? r.Moments[0].Name : "",
                    Moment2 = r.Moments.Count >= 2 ? r.Moments[1].Name : "",
                    Moment3 = r.Moments.Count >= 3 ? r.Moments[2].Name : "",
                    Moment4 = r.Moments.Count >= 4 ? r.Moments[3].Name : "",
                }).ToList();

            dataGridView2.DataSource = null;
            dataGridView2.DataSource = classesInGrid;
            dataGridView2.AutoResizeColumns();
            dataGridView2.Update();
            dataGridView2.Refresh();

            List<Deltagare> deltagare = readVaulters();
            var vcount = deltagare.Count;

            var result = deltagare
                .Select(r => new
                {
                    Name = r.Name,
                    Team = r.Klubb,
                    Class = r.Klass,
                    ClassName = classes.Single(k => k.Name == r.Klass).Description,
                    Horse = r.Hast,
                    Lunger = r.Linforare,
                    Internal_Id = r.Id
                }).ToList();

            dataGridView1.DataSource = null;
            dataGridView1.DataSource = result;
            tabControl1.TabPages[0].Text = " Vaulters (" + vcount + ")";
            dataGridView1.Update();
            dataGridView1.Refresh();
            dataGridView1.AutoResizeColumns();

            comboBox1.SelectedValueChanged -= comboBox1_SelectedValueChanged;
            comboBox1.DataSource = null;

            List<ClassSelect> sl = new List<ClassSelect>();
            var l = from c in classes
                    select new ClassSelect
                    {
                        Text = c.Name + " " + c.Description,
                        Value = c.Name
                    };

            comboBox1.DataSource = l.ToList();
            comboBox1.DisplayMember = "Text";

            comboBox1.SelectedValueChanged += comboBox1_SelectedValueChanged;

            dataGridView3.Update();
            dataGridView3.Refresh();
            dataGridView3.AutoResizeColumns();
        }

        /// <summary>
        /// Fake results by putting a value between 0 and 10 in the result cell of all sheets in Inbox
        /// </summary>
        private void doFake()
        {
            UpdateProgressBarHandler(0);
           
            UpdateProgressBarLabel("");

            DirectoryInfo dirinfo = new DirectoryInfo(fakeboxFolder);
            var files = dirinfo.EnumerateFiles("*.xls*").ToList();
            UpdateProgressBarMax(files.Count());

            if (files.Count() == 0)
            {
                showMessageBox("No result files available");
                return;
            }

            var file = new FileInfo(fakefile);

            if(file.Exists)
            {
                file.Delete();
            }

            UpdateProgressBarHandler(0);
            UpdateProgressBarLabel("");
            UpdateProgressBarMax(files.Count());

            UpdateProgressBarHandler(0);
            UpdateProgressBarLabel("");
            UpdateProgressBarMax(files.Count());


            // add a new worksheet to the empty workbook
            //Dictionary<int, List<string>> data = new Dictionary<int, List<string>>();
            ConcurrentDictionary<int, List<string>> data = new ConcurrentDictionary<int, List<string>>();

            int rownumber = 0;

            Boolean woody = bool.Parse(ConfigurationManager.AppSettings["woody"]);






            foreach (var f1 in files)
            {
                //Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 50 }, f1 =>
                //{

                //Interlocked.Increment(ref rownumber);
                rownumber++;

                //lock (lockObject)
                //{
                    UpdateProgressBarLabel("Faking # " + rownumber+ " " + f1.Name);
                //}

                try
                {

                    using (var wb = new XLWorkbook(f1.FullName))
                    {
                        // No need to put the worksheet inside a "using" block because
                        // the workbook will dispose of the sheets. The worksheet is not
                        // created inside a loop and the workbook's dispose is being
                        // called immediately after using the worksheet.
                        var ws = wb.Worksheets.SingleOrDefault(w => w.Visibility == XLWorksheetVisibility.Visible);

                        var rand = Math.Round(new Random().NextDouble() * 10, 3);
                        if (f1.FullName.Contains("_A_") && woody)
                        {
                            rand = 6.5;
                        }


                        var range = ws.NamedRange("result");
                        var refersTo = range.RefersTo;
                        var cellRange = ws.Range(refersTo);
                        var cells = cellRange.Cells();
                        cells.First().Value = rand;


                        //List<string> d = new List<string>();

                        //refersTo = ws.NamedRange("klass").RefersTo;
                        //var klassName = ws.Range(refersTo).Cells().First().Value.ToString();

                        //refersTo = ws.NamedRange("bord").RefersTo;
                        //var bord = ws.Range(refersTo).Cells().First().Value.ToString();

                        //refersTo = ws.NamedRange("moment").RefersTo;
                        //var moment = ws.Range(refersTo).Cells().First().Value.ToString();

                        //refersTo = ws.NamedRange("id").RefersTo;
                        //var id = ws.Range(refersTo).Cells().First().Value.ToString();

                        //d.AddRange(new List<string> { klassName, f1.Name, id, moment, bord, rand.ToString(CultureInfo.InvariantCulture) });
                        ////data.Add(rownumber, d);
                        //data.TryAdd(rownumber, d);
                        wb.Save();
                    }
                }
                catch (Exception e)
                {
                    var s = e.Message;
                    UpdateMessageTextBox($"Exception : {f1.Name} " + s);
                    //showMessageBox("Exception :" + s);
                }
                finally
                {
                    //lock (lockObject)
                    //{
                        UpdateProgressBarHandler(rownumber);
                        UpdateProgressBarLabel("Faked " + f1.Name);
                    //}
                
                }
            }

            UpdateMessageTextBox("Allt fake klart");


            using (var package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Fakes");
                foreach (KeyValuePair<int, List<string>> kvp in data)
                {
                    int row = (Int32)kvp.Key;

                    for (int i = 0; i < kvp.Value.Count; i++)
                    {
                        if (i == 0)
                        {
                            worksheet.Cells[row, i + 1].Value = kvp.Value[i];
                        }
                        else if (i == 5)
                        {
                            worksheet.Cells[row, i + 1].Value = float.Parse(kvp.Value[i],CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            worksheet.Cells[row, i + 1].Value = kvp.Value[i];
                        }

                    }
                }
                package.Save();
            }
            showMessageBox("Faking results completed, see " + Path.GetFileName(fakefile));
        }

        private delegate void UpdateProgressBarCallback(int barValue);
        private delegate void UpdateProgressBarLabelCallback(string text);
        private delegate void UpdateProgressBarMaxCallback(int barValue);
        private delegate void UpdateMessageTextBoxCallback(string text);

        private void UpdateProgressBarHandler(int barValue)
        {
            if (this.progressBar1.InvokeRequired)
                this.BeginInvoke(new UpdateProgressBarCallback(this.UpdateProgressBarHandler), new object[] { barValue });
            else
            {
                // change your bar
                this.progressBar1.Value = barValue;
                this.progressBar1.Refresh();
            }
        }

        private void UpdateProgressBarMax(int barValue)
        {
            if (this.progressBar1.InvokeRequired)
                this.BeginInvoke(new UpdateProgressBarMaxCallback(this.UpdateProgressBarMax), new object[] { barValue });
            else
            {
                // change your bar
                this.progressBar1.Maximum = barValue;
            }
        }

        private void UpdateProgressBarLabel(string text)
        {
            if (this.progressLabel.InvokeRequired)
                this.BeginInvoke(new UpdateProgressBarLabelCallback(this.UpdateProgressBarLabel), new object[] { text });
            else
            {
                // change your bar
                this.progressLabel.Text = text;
                this.progressLabel.Refresh();
            }
        }

        

        public void UpdateMessageTextBox(string text)
        {
            if (this.textBox1.InvokeRequired)
                this.BeginInvoke(new UpdateMessageTextBoxCallback(this.UpdateMessageTextBox), new object[] { text });
            else
            {
                // change your text
              this.textBox1.AppendText(text + System.Environment.NewLine);// (char)13);
            }
        }


        // Here are the background workers...
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            this.doFake();
        }

 

    
        // Populate the tables with classes and vaulters
        private void button4_Click(object sender, EventArgs e)
        {
            ReadClassesAndVaultersFromStartlist();


        }

        // Fake results in Inbox files
        private void buttonFakeResults_Click(object sender, EventArgs e)
        {
            backgroundWorkerFakeResults.RunWorkerAsync();
            bool hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerFakeResults.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(50);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                       //re-checking.
            }

        }

        private Color GetRowColor(int categoryNumber)
        {
            if ((categoryNumber) % 8 < 4)
                return Color.White; //default row color
            else
                return Color.LightGray; //alternate row color
        }

        private void LoadSortedResultsForClass(string className, string description)
        {
            if(!File.Exists(sortedresultsfile))
            {
                showMessageBox("No sorted file exists, copying result file and do initial sort!");
                doSort();

            }

            using (ExcelPackage p = new ExcelPackage(new FileInfo(sortedresultsfile)))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[className];
                if (ws != null)
                {
                    var data = ToDataTable(ws);
                    dataGridView3.DataSource = null;
                    dataGridView3.DataSource = data;
                    tabControl1.TabPages[2].Text = description;
                    dataGridView3.AutoResizeColumns();
                    dataGridView3.Columns[1].Visible = false;
                    dataGridView3.Columns[2].Visible = false;
                    dataGridView3.Columns[12].Visible = false;
                    dataGridView3.Columns[13].Visible = false;
                }
            }
        }

        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex == 0)
                return;

            if (e.RowIndex > 5)
            {
                int i = e.RowIndex;
                Color c = GetRowColor(i + 2);
                dataGridView3.Rows[i].DefaultCellStyle.BackColor = c;
            }
            else
            {
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Orange;
            }
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int ind = e.ColumnIndex;
            
            if (e.RowIndex > 5 && ind > 6 && ind < 11)
            {
                var curcell = dataGridView3[6, e.RowIndex];

                if (e.Value.ToString() != "")
                {
                    //e.CellStyle.BackColor = Color.White;

                }
                else if (curcell.Value.ToString().Length > 0)
                {
                    e.CellStyle.BackColor = Color.Green;
                }
                else if (curcell.Value.ToString().Length == 0)
                {
                    dataGridView3[11, e.RowIndex].Value = "";
                    dataGridView3[14, e.RowIndex].Value = "";
                }
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            ClassSelect sel = (sender as ComboBox).SelectedItem as ClassSelect;
            string value = sel.Value;
            string text = sel.Text;
            (sender as ComboBox).Refresh();
       
            if (value == null)
                return;

            LoadSortedResultsForClass(value, text);
            dataGridView3.AutoResizeColumns();
            dataGridView3.Update();
            dataGridView3.Refresh();
            (sender as ComboBox).SelectionLength = 0;
        }


    public Excel._Application printResultsExcelHandler(string className,string filename)
        {
            Excel.Application MyApp = null;
            Excel.Workbook MyBook = null;
            Excel.Workbooks workbooks = null;
            Excel.Worksheet MySheet = null;
            bool preliminiaryResults = checkBox1.Checked;
            string fullpath = Path.Combine(printedresultsFolder, filename);
            string pdfFullPath = fullpath + ".pdf";
            String noresults = ConfigurationManager.AppSettings["noresults"];

            List<String> noresultsList = noresults.Split(',').ToList();




            // Domare
            int counter = 0;
            Klass klass = readClasses().Find(c => c.Name.Equals(className));
            List<String> judgelist = new List<string>();
            //String totalJudge = "";
            foreach (Moment mom in klass.Moments)
            {
                counter++;
                string judgetext = string.Format("{0,15}", mom.Name);
                // Can be calculated, but not yet...
                int subcounter = 8;
               
                foreach (SubMoment submom in mom.SubMoments)
                {
             
                    if (submom.Table.judge.Fullname.Trim().Length > 0)
                    {
                        string judgeName = string.Format("{0,-30}", submom.Table.judge.Fullname);

                        judgetext = judgetext + "   " + submom.Table.Name + " : " +judgeName;
                    }
                     
                }
                judgelist.Add(judgetext);
               // totalJudge = totalJudge + judgetext + "\n";



                //foreach (SubMoment submom in mom.SubMoments)
                //{
                //    int row = classWorksheet.Cells[$"round{counter}"].Start.Row;
                //    classWorksheet.Cells[row, subcounter].Value = submom.Name;
                //    subcounter++;
                //}
            }

            try
            {
                MyApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false
                    
                    //DisplayAlerts = true
                };
                workbooks = MyApp.Workbooks;

                MyBook = workbooks.Open(sortedresultsfile, ReadOnly: true);
                MySheet = MyBook.Sheets[className];

                var usedRange = MySheet.UsedRange;


                if (noresultsList.Contains(className))
                {
                    var range = MySheet.get_Range("H7", "O50");
                    range.NumberFormat = ";;;";               
                }

                //MySheet.Activate();
                
                //if (checkBox1.Checked)
                //{
                //    MySheet.PageSetup.RightHeaderPicture.Filename = preliminaryResults;
                //}
                //else
                //{
                //    MySheet.PageSetup.RightHeaderPicture.Filename = logovoid;
                //}

                //MyApp.Visible = true;
                //string fullpath = Path.Combine(printedresultsFolder, filename);
                MySheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFullPath);
                //MySheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFullPath);




                //Excel.Range range2 = MySheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeAllFormatConditions);

                // foreach (Excel.Range c in range2.Cells)
                // {

                //     var color = c.DisplayFormat.Interior.Color;
                //     if(color==65280)
                //     {
                //         c.Interior.Color = 65280;
                //     }
                //     var t2 = c.Value2;
                //     var t1 = c.Value;

                // }





                //  MySheet.SaveAs(pdfFullPath + ".html", Excel.XlFileFormat.xlHtml);

                //Microsoft.Office.Interop.Excel.XlHtmlType t = Excel.XlHtmlType.xlHtmlCalc;

                //FileInfo excelfile = new FileInfo(sortedresultsfile);
                //String myName = MySheet.Name;

                //var WorksheetHtml = new ExcelToHtml.ToHtml(excelfile,myName);
                //WorksheetHtml.
                //string html = WorksheetHtml.GetHtml();



                //Excel.PublishObject p= MyBook.PublishObjects.Add(Excel.XlSourceType.xlSourceSheet, pdfFullPath + "AAA.html", MySheet.Name,null, t,MySheet.Name);
                //p.Publish();


                MyApp.DisplayAlerts = false;
                MyBook.Close();
                MyApp.DisplayAlerts = true;
                MyApp.Quit();



                var fileinfo = new FileInfo(sortedresultsfile);

                using (var pck = new ExcelPackage(fileinfo))
                {
                    ExcelWorksheet sheet = pck.Workbook.Worksheets[className];

                    var strUsedRange = sheet.Rows;
                    var dim =  sheet.Dimension;
                    String ddim = dim.ToString();
                    //var tableRange = sheet.Cells[strUsedRange];//.LoadFromDataTable(_dataTable, true, style);

                    var exporter = sheet.Cells[ddim].CreateHtmlExporter();
                    var settings = exporter.Settings;

                    settings.Culture = CultureInfo.InvariantCulture;
                    settings.TableId = "voltige-table";
                    settings.Accessibility.TableSettings.AriaLabel = "Voltige";
                    settings.SetColumnWidth = true;

                    // Export Html and CSS
                    exporter.Settings.Minify = true;
                    String Css = exporter.GetCssString();
                    String Html = exporter.GetHtmlString();

                    String finale = pdfFullPath + ".html";

                    File.Delete(finale);

                    List<String> JudegsList = judgelist.Select(p => "<li>" + p + "</li>").ToList();
                    String JudegsListString =  String.Join("", JudegsList);

                    String judlisthtml = @"<ul style=""list-style-type:none"">" + JudegsListString + "</ul>";

                   
                    String headerTable = @" <table border=""1"" width=""100 %"">
                                                   <tr>
                                                    <td> Voltige-SM <br> Billdal, 2022-07-13 -> 2022-07-16</td>
                                                    <td>" + judlisthtml + @"</td>
                                                    <td> Country </td>
                                                  </tr>
                                            </table>";

                    //String header =
                    //    "<header>" +
                    //    "Voltige-SM" +
                    //    "Billdal, 2022-07-13 -> 2022-07-16" + JudegsListString+
                    //    "</header>";

                    File.AppendAllText(finale, "<html>");
                    File.AppendAllText(finale, "<style>");
                    File.AppendAllText(finale, Css);
                    File.AppendAllText(finale, "</style>");
                    File.AppendAllText(finale, headerTable);

                    File.AppendAllText(finale, Html);
                    File.AppendAllText(finale, "</html>");
                  
                }

                createHtml(className);

            }
            catch(Exception e)
            {
              this.UpdateMessageTextBox($"Save to PDF failed for {className} : {e.Message}");
              showMessageBox(e.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(MySheet);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(workbooks);
                Marshal.FinalReleaseComObject(MyApp);
                MySheet = null;
                MyBook = null;
                workbooks = null;
                MyApp = null;
            }

      // Fix logos 

          try
          {
            
            //var sponsorlogo = Path.Combine(Form1.logosFolder, "sponsor.png");
            //var complogo = Path.Combine(Form1.logosFolder, "competition.png");
            var preliminary = Path.Combine(Form1.logosFolder, "preliminaryresults.png");
            var ridsport = Path.Combine(Form1.logosFolder, "logo_ridsport_top.png");
            var datelogo = Path.Combine(Form1.logosFolder, "date.png");
              var noresultlogo = Path.Combine(Form1.logosFolder, "nopoints.png");

                PdfDocument document = PdfReader.Open(pdfFullPath, PdfDocumentOpenMode.Modify);

            for (int i = 0; i < document.Pages.Count; ++i)
            {
              PdfPage page = document.Pages[i];

              // Make a layout rectangle.  
              //XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
              //using (XGraphics gfx = XGraphics.FromPdfPage(page))
              //{
              //  gfx.DrawString($" {now:F} -  Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);
              //}
              using (XGraphics gfx = XGraphics.FromPdfPage(page))
              {
                var xim = XImage.FromFile(ridsport);
                gfx.ScaleTransform(0.4);
                gfx.DrawImage(xim, new Point(120, 10));
              }

            //using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //{
            //    var xim = XImage.FromFile(complogo);
            //    gfx.ScaleTransform(0.15);
            //    gfx.DrawImage(xim, new Point(600, 10));
            //}

                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
              {
                var xim = XImage.FromFile(datelogo);
                gfx.ScaleTransform(0.35);
                gfx.DrawImage(xim, new Point(260, 30));
              }

              //using (XGraphics gfx = XGraphics.FromPdfPage(page))
              //{
              //  var xim = XImage.FromFile(sponsorlogo);
              //  gfx.ScaleTransform(0.3);
              //  gfx.DrawImage(xim, new Point(2000, 30));
              //}

              if (preliminiaryResults)
              {
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                  var xim = XImage.FromFile(preliminary);
                  gfx.ScaleTransform(0.5);
                  gfx.DrawImage(xim, new Point(1300, 140));
                }
              }


                if (noresultsList.Contains(className))
                {
                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        var xim = XImage.FromFile(noresultlogo);
                        gfx.ScaleTransform(0.8);
                        gfx.DrawImage(xim, new Point(500,10));
                    }
                  }

             }

            document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            document.Save(pdfFullPath);
          }
          catch (Exception logoException )
          {
            this.UpdateMessageTextBox($"Save to PDF failed for {className} : {logoException.Message}");
          }

          return null;
        }


        private void createHtml(String className)
        {

                var deltagare = readVaulters();
                var classes = readClasses();
                var max = deltagare.Count();
                             

                //UpdateProgressBarHandler(0);
                //UpdateProgressBarMax(deltagare.Count);
                //UpdateProgressBarLabel("");

                var resultat = new FileInfo(sortedresultsfile);

                //FileInfo resultat = new FileInfo(resultfile);

                // keep track of first vaulter / class so we know if we shall copy range or not 
                List<string> set = new List<string>();
                Dictionary<string, int> vaulterInClassCounter = new Dictionary<string, int>();
                foreach (Klass c in classes)
                {
                    vaulterInClassCounter[c.Name] = 0;
                }


                ExcelRange toRange;
                ExcelRange fromRange;

                using (var results = new ExcelPackage(resultat))
                {

                String klassnamn = className;
                //foreach (Klass klass in classes)
                //{
                Klass klass = classes.First(c => c.Name.Equals(klassnamn));
                String file = null; 
                String text = null;

                String file2 = null;
                String resultatheadertext = null;

                String _file3 = null;
                String _text3 = null;

                String _file4 = null;
                String _text4 = null;

                //String cssfile = null;
                String headfile = null;
                String head = null;

                int moments = klass.Moments.Count();
                List<Judge> judges = klass.Moments[0].SubMoments.Select(s => s.Table.judge).ToList();

                _file4 = Path.Combine(Environment.CurrentDirectory, "html/mallMain.html");
                headfile= Path.Combine(Environment.CurrentDirectory, "html/HTML_head.html");
                //cssfile = Path.Combine(Environment.CurrentDirectory, "html/stylesheet.css");

                if (klass.ResultTemplate.Equals("GK2"))
                {
                    file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare2moment.html");
                    file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header2domare2moment.html");
                    _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat2domare2moment.html");
                } else if (klass.ResultTemplate.Equals("ResultTemplate"))
                   {
                    if (moments == 3)
                    {
                        file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top4domare3moment.html");
                        file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4domare3moment.html");
                        _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4domare3moment.html");
                    }
                    else if (moments == 4)
                    {
                        file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top4domare4moment.html");
                        file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4domare4moment.html");
                        _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4domare4moment.html");
                    }
                }

                head= File.ReadAllText(headfile);
                text = File.ReadAllText(file);
                resultatheadertext = File.ReadAllText(file2);
                _text3 = File.ReadAllText(_file3);
                _text4 = File.ReadAllText(_file4);


                var sheet = results.Workbook.Worksheets[klass.Name];



                text = text.Replace("{KLASS}", "Klass" + klass.Name + " - "+ klass.Description);
                int counter = 0;
                    foreach (Moment moment in klass.Moments)
                    {
                        counter++;
                        text = text.Replace("{MOMENT_"+counter+"}",moment.Name); 
                        resultatheadertext = resultatheadertext.Replace("{MOMENT_" + counter + "}", moment.Name);

                    foreach (SubMoment submoment in moment.SubMoments)
                     {
                            String table = submoment.Table.Name;
                            String judgename = submoment.Table.judge.Fullname;
                            text = text.Replace("{MOMENT_" + counter + "_DOMARE_"+table+"}", judgename);
                            resultatheadertext = resultatheadertext.Replace("{MOMENT_" + counter + "_" + table + "}", submoment.Name);
                     }
                    }

                File.WriteAllText("test.html", text);
                File.WriteAllText("test2.html", resultatheadertext);
                
                int rowbase = 7;
                int endrow = sheet.Dimension.End.Row;

                String textrows = "";

                String noresults = ConfigurationManager.AppSettings["noresults"];
                List<String> noresultsList = noresults.Split(',').ToList();
                Boolean noresultsInClass = noresultsList.Contains(klassnamn);

                int currentRowInTable = 0;

                for (int row = rowbase; row < endrow; row += 4)
                {
                    currentRowInTable++;

                    _text3 = File.ReadAllText(_file3);
                    String text3 = _text3;


                    toRange = sheet.Cells[row, 1, row + 3, 15];

                    String placering = toRange[row + 1, 1].GetValue<String>();// 

                    String name = toRange[row + 1, 4].GetValue<String>();// = d.Name;
                    String linforare = toRange[row + 2, 4].GetValue<String>();// = d.Name;

                    String club = toRange[row + 1, 6].GetValue<String>();// = d.Klubb;
                    String horse = toRange[row + 2, 6].GetValue<String>();// = d.Hast;

                    String tot = toRange[row + 1, 15].Text; // GetValue<String>();// 


                    if (noresultsInClass) tot = "-";


                    text3 = text3.Replace("{PLACERING}", placering);
                    text3 = text3.Replace("{NAMN}", name);
                    text3 = text3.Replace("{KLUBB}", club);
                    text3 = text3.Replace("{LINFORARE}", linforare);
                    text3 = text3.Replace("{HAST}", horse);
                    text3 = text3.Replace("{TOT}", tot);

                    counter = 0;
                    foreach (Moment moment in klass.Moments)
                    {
                        int rowindex = counter + 1;
                        String mom = toRange[row + counter, 7].GetValue<String>();// = moment;
                        text3 = text3.Replace("{MOMENT_" + rowindex + "}", moment.Name);

                        String momsum = toRange[row + counter, 12].Text;// GetValue<String>();

                        if (noresultsInClass) momsum = "-";

                        text3 = text3.Replace("{MOMENTSUM_" + rowindex + "}", momsum);

                        var tt = toRange[row, 1, row, 15];

                        int subcounter = 0;
                        foreach (SubMoment submoment in moment.SubMoments)
                        {
                            subcounter++;
                            String table = submoment.Table.Name;
                            String point = toRange[row + counter, 7 + subcounter].Text; // GetValue<String>();// = moment;;
                            String key = "{POANG_" + rowindex + "_" + table + "}";
                            String keycell = "{POANG_" + rowindex + "_" + table + "_CLASS}";

                            if (point == "")
                            {
                                text3 = text3.Replace(keycell, "empty");
                            }
                            else
                            {
                                text3 = text3.Replace(keycell, "");
                            }

                            if (noresultsInClass) point = "-";
                   
                            text3 = text3.Replace(key, point);
                 


                        }
                        counter++;
                    }
                    textrows = textrows + text3;
                }

                File.WriteAllText("test3.html", textrows);


                // Skapa fil
                _text4=_text4.Replace("{HEAD}", head);
                _text4=_text4.Replace("{TOP}", text);
                _text4=_text4.Replace("{HEADER}", resultatheadertext);
                _text4=_text4.Replace("{DATA}", textrows);
                File.WriteAllText(Path.Combine(htmlResultsFolder, klass.Name + " - " + klass.Description +".html"), _text4);


                int h = 5;

              //  String club = sheet.Cells[row, 6].GetValue<String>();// = d.Klubb;
              //  String horse = sheet.Cells[row +1, 6].GetValue<String>();// = d.Hast;



                //foreach ()

                //        // We have more than 1 competitor in the class and need to copy the ekipage range
                //        //if (set.Contains(klass))
                //        if (vaulterInClassCounter[klass] > 0)
                //        {
                //            toRange = sheet.Cells[row + 1, 1, row + 4, fromRange.End.Column];
                //            fromRange.Copy(toRange);




                // }



                //int deltagarCounter = 0;

                //    foreach (Deltagare d in deltagare)
                //    {
                //        deltagarCounter++;
                //        var klass = d.Klass;



                //        var sheet = results.Workbook.Worksheets[klass];

                //        int row = sheet.Dimension.End.Row;

                //        fromRange = sheet.Cells["ekipage"];


                //        // We have more than 1 competitor in the class and need to copy the ekipage range
                //        //if (set.Contains(klass))
                //        if (vaulterInClassCounter[klass] > 0)
                //        {
                //            toRange = sheet.Cells[row + 1, 1, row + 4, fromRange.End.Column];
                //            fromRange.Copy(toRange);


                //            //Set formatting
                //            for (int i = 1; i < 5; i++)
                //            {
                //                var theclass = classes.Single(c => c.Name == klass);
                //                var endcol = 11;

                //                if (theclass.ResultTemplate.Trim().EndsWith("1"))
                //                {
                //                    endcol = 8;
                //                    ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                //                    _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1_2.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                //                    _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                //                }

                //                if (theclass.ResultTemplate.Trim().EndsWith("1B"))
                //                {
                //                    endcol = 9;
                //                    ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 9, row + i, endcol);
                //                    var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                //                    _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1_2.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 9, row + i, endcol);
                //                    var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                //                    _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                //                }

                //                if (theclass.ResultTemplate.Trim().EndsWith("2"))
                //                {
                //                    endcol = 9;
                //                    ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                //                    _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1_2.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                //                    _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                //                }

                //                if (theclass.ResultTemplate.Trim().EndsWith("M3"))
                //                {
                //                    endcol = 10;
                //                    ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                //                    _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1_2.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                //                    _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                //                }


                //                if (theclass.ResultTemplate.Trim().EndsWith("K3"))
                //                {
                //                    endcol = 10;

                //                    ExcelAddress _formatRangeAddress = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond1 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress);
                //                    _cond1.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2);
                //                    _cond2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2.Style.Fill.BackgroundColor.Index = 3;

                //                    endcol = 11;
                //                    ExcelAddress _formatRangeAddressB = new ExcelAddress(row + i, endcol, row + i, endcol);
                //                    var _cond1B = sheet.ConditionalFormatting.AddExpression(_formatRangeAddressB);
                //                    _cond1B.Formula = $"COUNTBLANK($G{row + i})=0";
                //                    _cond1B.StopIfTrue = false;
                //                    _cond1B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    var color = System.Drawing.ColorTranslator.FromHtml("#E2EBD5");
                //                    //_cond1B.Style.Fill.BackgroundColor.Index = -4142;
                //                    _cond1B.Style.Fill.BackgroundColor.Color = color;

                //                    //ExcelAddress _formatRangeAddress2B = new ExcelAddress(row + i, endcol, row + i, endcol);
                //                    //var _cond2B = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2B);
                //                    //_cond2B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    //_cond2B.Style.Fill.BackgroundColor.Index = 35;

                //                }

                //                if (theclass.ResultTemplate.Trim().EndsWith("ResultTemplate"))
                //                {
                //                    endcol = 11;

                //                    ExcelAddress _formatRangeAddress = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond1 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress);
                //                    _cond1.Formula = $"COUNTBLANK($G{row + i})=1";
                //                    _cond1.StopIfTrue = true;

                //                    ExcelAddress _formatRangeAddress2 = new ExcelAddress(row + i, 8, row + i, endcol);
                //                    var _cond2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2);
                //                    _cond2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                    _cond2.Style.Fill.BackgroundColor.Index = 3;

                //                }

                //            }
                //        }
                //        else
                //        {
                //            // First competitor can use predefined "ekipage" range that exists on sheet
                //            set.Add(klass);
                //            string adress = fromRange.Address;
                //            toRange = fromRange;

                //        }

                //        // Names, horses etc
                //        var startrow = toRange.Start.Row;
                //        sheet.Cells[startrow + 1, 4].Value = d.Name;
                //        sheet.Cells[startrow + 2, 4].Value = d.Linforare;
                //        sheet.Cells[startrow + 1, 6].Value = d.Klubb;
                //        sheet.Cells[startrow + 2, 6].Value = d.Hast;

                //        // The Id of the ekipage
                //        sheet.Cells[startrow, 2].Value = d.Id;
                //        sheet.Cells[startrow + 1, 2].Value = d.Id;
                //        sheet.Cells[startrow + 2, 2].Value = d.Id;
                //        sheet.Cells[startrow + 3, 2].Value = d.Id;

                //        // We need the klass to get the Table name
                //        var tklass = classes.Single(c => c.Name == klass);

                //        startrow = toRange.Start.Row;
                //        int momentIndex = 0; // ID
                //        foreach (Moment moment in tklass.Moments)
                //        {
                //            // ID generation
                //            var colnum = 8;
                //            momentIndex++;
                //            foreach (SubMoment submoment in moment.SubMoments)
                //            {
                //                // ID
                //                //string id = d.Id + "_" + klass + "_" + moment.Name.Replace(' ', '_') + "_" + submoment.Table.Name;
                //                string id = d.Id + "_" + momentIndex + "_" + submoment.Table.Name; //ID

                //                var rng = sheet.Cells[startrow, colnum];
                //                sheet.Names.Add(id, rng);
                //                colnum++;
                //            }
                //            startrow++;
                //        }
                //        UpdateProgressBarHandler(deltagarCounter);
                //        UpdateProgressBarLabel("Added " + d.Name);
                //        vaulterInClassCounter[klass]++;

                //        var modCounter = vaulterInClassCounter[klass] % 9;

                //        // Pagebreak every 9 vaulter / class
                //        if (modCounter == 0)
                //        {
                //            int lastRow = toRange.End.Row;
                //            sheet.Row(lastRow).PageBreak = true;
                //        }
                //    }
                //    results.Save();
            }

            UpdateProgressBarLabel("All vaulters added to result file");
 
        }


        //public Excel._Application printResultsExcelHandler2(string className, string filename)
        //{
        //    Excel.Application MyApp = null;
        //    Excel.Workbook MyBook = null;
        //    Excel.Workbooks workbooks = null;
        //    Excel.Worksheet MySheet = null;
        //    bool preliminiaryResults = checkBox1.Checked;
        //    string fullpath = Path.Combine(printedresultsFolder, filename);
        //    string pdfFullPath = fullpath + ".pdf";
        //    String noresults = ConfigurationManager.AppSettings["noresults"];

        //    List<String> noresultsList = noresults.Split(',').ToList();

        //    try
        //    {
        //        MyApp = new Excel.Application
        //        {
        //            Visible = false,
        //            ScreenUpdating = false

        //            //DisplayAlerts = true
        //        };
        //        workbooks = MyApp.Workbooks;
        //        MyBook = workbooks.Open(sortedresultsfile, ReadOnly: true);
        //        MySheet = MyBook.Sheets[className];



        //        if (noresultsList.Contains(className))
        //        {
        //            var range = MySheet.get_Range("H7", "O50");
        //            range.NumberFormat = ";;;";
        //        }

        //        //MySheet.Activate();

        //        //if (checkBox1.Checked)
        //        //{
        //        //    MySheet.PageSetup.RightHeaderPicture.Filename = preliminaryResults;
        //        //}
        //        //else
        //        //{
        //        //    MySheet.PageSetup.RightHeaderPicture.Filename = logovoid;
        //        //}

        //        //MyApp.Visible = true;
        //        //string fullpath = Path.Combine(printedresultsFolder, filename);
        //        MySheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFullPath);

        //        MyApp.DisplayAlerts = false;
        //        MyBook.Close();
        //        MyApp.DisplayAlerts = true;
        //        MyApp.Quit();
        //    }
        //    catch (Exception e)
        //    {
        //        this.UpdateMessageTextBox($"Save to PDF failed for {className} : {e.Message}");
        //        showMessageBox(e.Message);
        //    }
        //    finally
        //    {
        //        Marshal.FinalReleaseComObject(MySheet);
        //        Marshal.FinalReleaseComObject(MyBook);
        //        Marshal.FinalReleaseComObject(workbooks);
        //        Marshal.FinalReleaseComObject(MyApp);
        //        MySheet = null;
        //        MyBook = null;
        //        workbooks = null;
        //        MyApp = null;
        //    }

        //    // Fix logos 

        //    try
        //    {

        //        //var sponsorlogo = Path.Combine(Form1.logosFolder, "sponsor.png");
        //        //var complogo = Path.Combine(Form1.logosFolder, "competition.png");
        //        var preliminary = Path.Combine(Form1.logosFolder, "preliminaryresults.png");
        //        var ridsport = Path.Combine(Form1.logosFolder, "logo_ridsport_top.png");
        //        var datelogo = Path.Combine(Form1.logosFolder, "date.png");
        //        var noresultlogo = Path.Combine(Form1.logosFolder, "nopoints.png");

        //        PdfDocument document = PdfReader.Open(pdfFullPath, PdfDocumentOpenMode.Modify);

        //        for (int i = 0; i < document.Pages.Count; ++i)
        //        {
        //            PdfPage page = document.Pages[i];

        //            // Make a layout rectangle.  
        //            //XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
        //            //using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //            //{
        //            //  gfx.DrawString($" {now:F} -  Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);
        //            //}
        //            using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //            {
        //                var xim = XImage.FromFile(ridsport);
        //                gfx.ScaleTransform(0.4);
        //                gfx.DrawImage(xim, new Point(120, 10));
        //            }

        //            //using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //            //{
        //            //    var xim = XImage.FromFile(complogo);
        //            //    gfx.ScaleTransform(0.15);
        //            //    gfx.DrawImage(xim, new Point(600, 10));
        //            //}

        //            using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //            {
        //                var xim = XImage.FromFile(datelogo);
        //                gfx.ScaleTransform(0.35);
        //                gfx.DrawImage(xim, new Point(260, 30));
        //            }

        //            //using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //            //{
        //            //  var xim = XImage.FromFile(sponsorlogo);
        //            //  gfx.ScaleTransform(0.3);
        //            //  gfx.DrawImage(xim, new Point(2000, 30));
        //            //}

        //            if (preliminiaryResults)
        //            {
        //                using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //                {
        //                    var xim = XImage.FromFile(preliminary);
        //                    gfx.ScaleTransform(0.5);
        //                    gfx.DrawImage(xim, new Point(1300, 140));
        //                }
        //            }


        //            if (noresultsList.Contains(className))
        //            {
        //                using (XGraphics gfx = XGraphics.FromPdfPage(page))
        //                {
        //                    var xim = XImage.FromFile(noresultlogo);
        //                    gfx.ScaleTransform(0.8);
        //                    gfx.DrawImage(xim, new Point(500, 10));
        //                }
        //            }

        //        }

        //        document.Options.CompressContentStreams = true;
        //        document.Options.NoCompression = false;
        //        document.Save(pdfFullPath);
        //    }
        //    catch (Exception logoException)
        //    {
        //        this.UpdateMessageTextBox($"Save to PDF failed for {className} : {logoException.Message}");
        //    }

        //    return null;
        //}

        // Export Results for class
        private void printResults(string className, string description)
        {
          try
          {
            UpdateMessageTextBox($"Saving class '{className}' to PDF");
            printResultsExcelHandler(className, description);
            UpdateMessageTextBox($"Saving class '{className}' to PDF done...");
      }
          catch (Exception ee)
          {
            UpdateMessageTextBox($"Saving class {className} to PDF failed...");
            UpdateMessageTextBox(ee.Message);
      }

          GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
    }

        // Export Results for selected class
        private void button2_Click_1(object sender, EventArgs e)
        {
            ClassSelect sel = comboBox1.SelectedItem as ClassSelect ;
            string value = sel.Value;
            string text = sel.Text;
            printResults(value, text);
        }

        // Export Results for all classes
        private void button5_Click(object sender, EventArgs e)
        {
            var allClasses = readClasses();
            foreach (var cl in allClasses)
            {   
                printResults(cl.Name, cl.Name+" "+cl.Description);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void backgroundWorker5_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

        }

        private void backgroundWorker5_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }




        // Clear messages 
        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void dataGridView3_DataSourceChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
          try
          {
            UpdateMessageTextBox("Merging PDFs...");
            pdf.Merge(printedresultsFolder);
            UpdateMessageTextBox("Merging PDFs done...");
          }
          catch (Exception ee)
          {
            UpdateMessageTextBox("Failed to Merge PDFs ...");
            UpdateMessageTextBox(ee.Message);
          }

          try
          {
            UpdateMessageTextBox("Publishing results...");
            PDFtoHTML.GenerateHTML();
            UpdateMessageTextBox("Publish done...");
          }
          catch(Exception ee)
          {
            UpdateMessageTextBox("Failed to Publish ...");
            UpdateMessageTextBox(ee.Message);
          }
    }


      public class Horse : IComparable<Horse>
      {
        public class HorsePoint
        {
          public float thePoint;
          public string vaulter;
          public string klass;
        }

      public string Name;
        public float Max
        {
          get
          {
            return Points.Count > 0 ? Points.Max() : (float)0.0;
          }
        }

        public Horse()
        {
          Points = new List<float>();
        }


        public float Average
        {
          get
          {
            return Points.Count > 0 ? Points.Average() : (float)0.0;
          }
        }

        public List<float> Points;


        public int CompareTo(Horse other)
        {
          if (this.Average < other.Average) return 1;
          else if (this.Average > other.Average) return -1;
          else return 0;
        }
      }

      public class HPclass
      {
        public string Id;
        public string Name;
        public string Klass;
        public float point;

        public bool IsSM => !Klass.Contains(".");
        public bool IsNM => Klass.Contains(".1");
        public bool IsSMNM => Id.Contains(".2"); 

      public static HPclass Create(string hpline)
        {
          var data = hpline.Split(';');
          var hp = new HPclass
          {
            Id = data[0].Trim(), Name = data[1].Trim(), Klass = data[2].Trim(), point = float.Parse(data[3].Trim().Replace(",","."),CultureInfo.InvariantCulture)
          };
          return hp;
        }
      }

      public void CalculateHorsePoints2()
      {

      UpdateMessageTextBox($"Analyzing Horse points...");

      var teamclasses = ConfigurationManager.AppSettings["teamclasses"].Split(',').Select(s => s.Trim());
      var horsepointclasses = ConfigurationManager.AppSettings["horsepointclasses"].Split(',').Select(s => s.Trim());
      var horsepoints = Form1.horseresultfile;
      var horsepointsCalculated = Path.Combine(Form1.horseResultsFolder,"CalculatedHorsePoints.xlsx");
      var horsepointsCalculatedTemplate = Path.Combine(Application.StartupPath, "CalculatedHorsePoints_template.xlsx");


        var allHPs = File.ReadAllLines(horsepoints).Distinct().Select(HPclass.Create).ToList();
        var removedhorsepoints = allHPs.RemoveAll(hp => !horsepointclasses.Contains(hp.Klass));
        UpdateMessageTextBox($"Removed {removedhorsepoints} from calculation");

        
        //var allPointsInd  = allHPs.Where(hp => !teamclasses.Contains(hp.Klass));
        //var allPointsTeam = allHPs.Where(hp => teamclasses.Contains(hp.Klass));



       File.Delete(horsepointsCalculated);
       File.Copy(horsepointsCalculatedTemplate,horsepointsCalculated,true);

      var allSMNMPoints = allHPs.Where(hp => !(hp.IsSMNM && hp.IsNM));
      var allSMNMPointsInd = allSMNMPoints.Where(hp => !teamclasses.Contains(hp.Klass));
      var allSMNMPointsTeam = allSMNMPoints.Where(hp => teamclasses.Contains(hp.Klass));

        var allSMPoints = allHPs.Where(hp => hp.IsSM);
        var allSMPointsInd = allSMPoints.Where(hp => !teamclasses.Contains(hp.Klass));
        var allSMPointsTeam = allSMPoints.Where(hp => teamclasses.Contains(hp.Klass));

        var allNMPoints = allHPs.Where(hp => hp.IsNM);
        var allNMPointsInd = allNMPoints.Where(hp => !teamclasses.Contains(hp.Klass));
        var allNMPointsTeam = allNMPoints.Where(hp => teamclasses.Contains(hp.Klass));


      /*
       * Max SM/NM (Ind+Team)	Mean SM/NM (Ind + Team)	Max SM/NM (Ind)	Mean SM/NM (Ind)	Max SM/NM (Team)	Mean SM/NM (Team)
       */

        try
        {
          var horsepointGroup = from so in allHPs
            group so by so.Name
            into AllHorsePoints
            select new
            {
              HorseName = AllHorsePoints.Key,
              SMNMMax     = allSMNMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Max(s => s.point),
              SMNMAverage = allSMNMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Average(s => s.point),
              SMNMMaxInd = allSMNMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMNMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Max(s => s.point) : 0,
              SMNMMeanInd = allSMNMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMNMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Average(s => s.point) : 0,
              SMNMMaxTeam = allSMNMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMNMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Max(s => s.point) : 0,
              SMNMMeanTeam = allSMNMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMNMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Average(s => s.point) : 0,

              SMMax = allSMPoints.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              SMAverage = allSMPoints.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,
              SMMaxInd = allSMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              SMMeanInd = allSMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,
              SMMaxTeam = allSMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              SMMeanTeam = allSMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allSMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,

              NMMax = allNMPoints.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              NMAverage = allNMPoints.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,
              NMMaxInd = allNMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              NMMeanInd = allNMPointsInd.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPointsInd.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,
              NMMaxTeam = allNMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Max(hp => hp.point) : 0,
              NMMeanTeam = allNMPointsTeam.Any(hp => hp.Name == AllHorsePoints.Key) ? allNMPointsTeam.Where(hp => hp.Name == AllHorsePoints.Key).Average(hp => hp.point) : 0,
            };


          var fileinfo = new FileInfo(horsepointsCalculated);
          using (var results = new ExcelPackage(fileinfo))
          {
            var ws = results.Workbook.Worksheets[0];
            var row = 2;
            foreach (var horse in horsepointGroup)
            {

              ws.Cells[row, 1].Value = horse.HorseName;

            ws.Cells[row, 2].Value = horse.SMNMMax;
            ws.Cells[row, 3].Value = horse.SMNMAverage;
            ws.Cells[row, 4].Value = horse.SMNMMaxInd;
            ws.Cells[row, 5].Value = horse.SMNMMeanInd;
            ws.Cells[row, 6].Value = horse.SMNMMaxTeam;
            ws.Cells[row, 7].Value = horse.SMNMMeanTeam;

            //ws.Cells[row, 8].Value = horse.SMMax;
            //ws.Cells[row, 9].Value = horse.SMAverage;
            //ws.Cells[row, 10].Value = horse.SMMaxInd;
            //ws.Cells[row, 11].Value = horse.SMMeanInd;
            //ws.Cells[row, 12].Value = horse.SMMaxTeam;
            //ws.Cells[row, 13].Value = horse.SMMeanTeam;

            //ws.Cells[row, 14].Value = horse.NMMax;
            //ws.Cells[row, 15].Value = horse.NMAverage;
            //ws.Cells[row, 16].Value = horse.NMMaxInd;
            //ws.Cells[row, 17].Value = horse.NMMeanInd;
            //ws.Cells[row, 18].Value = horse.NMMaxTeam;
            //ws.Cells[row, 19].Value = horse.NMMeanTeam;
            row++;
            }

            results.Save();

          }
          UpdateMessageTextBox($"Horse points analyzed");
      }
        catch(Exception e)
        {
           UpdateMessageTextBox($"Failed to analyze horse points : {e.Message}");
        }

      }


 

    public void CalculateHorsePoints()
    {
      var teamclasses = ConfigurationManager.AppSettings["teamclasses"].Split(',').Select(s=>s.Trim());

      UpdateMessageTextBox($"Starting horse point calculation");
      var resultfile = Form1.sortedresultsfile;
      var horsefile = horseresultfile;

      if (!File.Exists(resultfile))
      {
        UpdateMessageTextBox($"{resultfile} not found, aborting horse point calculation");
        return;
      }

      FileInfo resultat = new FileInfo(resultfile);
      FileInfo horsefileInfo = new FileInfo(horsefile);

      List<string> horses = new List<string>();

      List<Horse> definedHorses = new List<Horse>();
      List<Horse> definedHorsesTeam = new List<Horse>();
      List<Horse> definedHorsesInd = new List<Horse>();

      var classes = readClasses();

      using (ExcelPackage results = new ExcelPackage(resultat))
      {
        try
        {
          foreach (var cl in classes)
          {
            UpdateMessageTextBox($"Getting horse points from class {cl.Name} - {cl.Description}");
            int startRow = 7;
            ExcelWorksheet ws = results.Workbook.Worksheets[cl.Name];
            var maxrow = ws.Dimension.End.Row;

            int ekipages = (maxrow - startRow + 1) / 4;

            for (int ekipage = 0; ekipage < ekipages; ekipage++)
            {
              var currentStartRow = startRow + (ekipage * 4);
              var horsename = ws.Cells[currentStartRow + 2, 6].Value.ToString();
              horses.Add(horsename);

              if (!definedHorses.Any(h => h.Name == horsename))
              {
                Horse h1 = new Horse();
                h1.Name = horsename;
                definedHorses.Add(h1);
              }

              
              // TEAM
              if (teamclasses.Contains(ws.Name)) 
              {
                if (!definedHorsesTeam.Any(h => h.Name == horsename))
                {
                  Horse h1 = new Horse();
                  h1.Name = horsename;
                  definedHorsesTeam.Add(h1);
                }
              }
              // IND
              else
              {
                if (!definedHorsesInd.Any(h => h.Name == horsename))
                {
                  Horse h1 = new Horse();
                  h1.Name = horsename;
                  definedHorsesInd.Add(h1);
                }
              }

              var curhorse = definedHorses.Single(h => h.Name == horsename);

              for (int arow = 0; arow < 4; arow++)
              {
                var momenttext = ws.Cells[currentStartRow + arow, 7].Value.ToString();
                if (momenttext.Length > 1) // we may have points
                {
                  var point = ws.Cells[currentStartRow + arow, 8].GetValue<float>();
                  if (point > 0)
                  {
                    curhorse.Points.Add(point);

                    // TEAM
                    if (teamclasses.Contains(ws.Name))
                    {
                      var curhorse1 = definedHorsesTeam.Single(h => h.Name == horsename);
                      curhorse1.Points.Add(point);

                    }
                    // IND
                    else
                    {
                      var curhorse2 = definedHorsesInd.Single(h => h.Name == horsename);
                      curhorse2.Points.Add(point);
                    }
                  }
                }
              }
            }
          }

        }
        catch (Exception ex)
        {
          var str = ex.Message;
          UpdateMessageTextBox(str);
        }
        finally
        {

        }
      }

      UpdateMessageTextBox($"Getting horse points from all classes done");
      var all = horses.Distinct().ToList();
      all.RemoveAll(s => s == "A4");
      definedHorses.RemoveAll(h => h.Name == "A4");
      definedHorsesTeam.RemoveAll(h => h.Name == "A4");
      definedHorsesInd.RemoveAll(h => h.Name == "A4");

      definedHorses.Sort();
      definedHorsesTeam.Sort();
      definedHorsesInd.Sort();

      File.Delete(horsefileInfo.FullName);

      using (ExcelPackage results = new ExcelPackage(horsefileInfo))
      {
        try
        {
          var sheet = results.Workbook.Worksheets.Add("Horse points team+ind");
          var sheet2 = results.Workbook.Worksheets.Add("Horse points team");
          var sheet3 = results.Workbook.Worksheets.Add("Horse points ind");
          sheet.Cells.Style.Numberformat.Format = @"0.000";
          sheet.Cells[1, 1].Value = "Häst";
          sheet.Cells[1, 3].Value = "Högsta enskilda poäng";
          sheet.Cells[1, 2].Value = "Medelpoäng";
          sheet.Cells[1, 4].Value = "Samtliga poäng";

          sheet2.Cells.Style.Numberformat.Format = @"0.000";
          sheet2.Cells[1, 1].Value = "Häst";
          sheet2.Cells[1, 3].Value = "Högsta enskilda poäng";
          sheet2.Cells[1, 2].Value = "Medelpoäng";
          sheet2.Cells[1, 4].Value = "Samtliga poäng";

          sheet3.Cells.Style.Numberformat.Format = @"0.000";
          sheet3.Cells[1, 1].Value = "Häst";
          sheet3.Cells[1, 3].Value = "Högsta enskilda poäng";
          sheet3.Cells[1, 2].Value = "Medelpoäng";
          sheet3.Cells[1, 4].Value = "Samtliga poäng";

          int row = 1;

          foreach (Horse h in definedHorses)
          {
            row = row + 1;
            sheet.Cells[row, 1].Value = h.Name;
            sheet.Cells[row, 3].Value = h.Max;
            sheet.Cells[row, 2].Value = h.Average;
            for (int i = 0; i < h.Points.Count; i++)
            {
              sheet.Cells[row, 4 + i].Value = h.Points[i];
            }

          }
          sheet.Cells.AutoFitColumns();

          row = 1;
          foreach (Horse h in definedHorsesTeam)
          {
            row = row + 1;
            sheet2.Cells[row, 1].Value = h.Name;
            sheet2.Cells[row, 3].Value = h.Max;
            sheet2.Cells[row, 2].Value = h.Average;
            for (int i = 0; i < h.Points.Count; i++)
            {
              sheet2.Cells[row, 4 + i].Value = h.Points[i];
            }

          }
          sheet2.Cells.AutoFitColumns();

          row = 1;
          foreach (Horse h in definedHorsesInd)
          {
            row = row + 1;
            sheet3.Cells[row, 1].Value = h.Name;
            sheet3.Cells[row, 3].Value = h.Max;
            sheet3.Cells[row, 2].Value = h.Average;
            for (int i = 0; i < h.Points.Count; i++)
            {
              sheet3.Cells[row, 4 + i].Value = h.Points[i];
            }

          }
          sheet3.Cells.AutoFitColumns();
          UpdateMessageTextBox($"{horsefile} created ! ");
        }
        catch (Exception ex)
        {
          UpdateMessageTextBox($"Horse point Error! ");
          UpdateMessageTextBox(ex.Message);
        }
        finally
        {
          results.Save();
          UpdateMessageTextBox($"{horsefile} saves ! ");
        }
      }
    
    }

    private void button3_Click(object sender, EventArgs e)
    {
      try
      {
        CalculateHorsePoints2();
      }
      catch (Exception ee)
      {
        UpdateMessageTextBox("Error in Horse point calc!");
        UpdateMessageTextBox(ee.Message);

      }
    }

        private void button6_Click(object sender, EventArgs e)
        {
            extractFromSortedFile();
        }
    }
    public static class Extension
    {
        public static bool IsNumeric(this string s)
        {
            float output;
            return float.TryParse(s, out output);
        }
    }
}
