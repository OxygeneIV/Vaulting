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
using System.Diagnostics;
using System.Globalization;
using System.Net.NetworkInformation;
using ClosedXML.Excel;
//using RasterEdge.Imaging.Basic;
//using RasterEdge.XDoc.Excel;

//using Spire.Xls;


namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        private static string root;
        private static string resultfile;
        private static string startlistfile;
        private static string inbox;
        private static string outbox;
        private static string fakebox;
        private static string backup;
        private static bool showmessageboxes;
        private static bool competitionstarted;


        private static string sortedresultsfile;
       public static string printedresults;
       public static string mergedresults;
      public static string publishresults;
    private static string workingDirectory;
        private static string logo;
        private static string logovoid;
        private static string preliminaryResults;
        private static bool fake;
        private static string fakefile;




        public Form1()
        {
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

                if (!fake)
                    buttonFakeResults.Enabled = false;

                if(competitionstarted)
                {
                    buttonPopulateSheetsWithVaulters.Enabled = false;
                    buttonCreateResultSheets.Enabled = false;
                }

                // Folders
                List<string> foldersToCreate = new List<string>();

                inbox = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["inbox"]);
                foldersToCreate.Add(inbox);

                fakebox = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["fakebox"]);
                foldersToCreate.Add(fakebox);

                backup = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["backup"]);
                foldersToCreate.Add(backup);

                outbox = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["outbox"]);
                foldersToCreate.Add(outbox);

                printedresults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["printedresults"]);
                foldersToCreate.Add(printedresults);

                 mergedresults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["mergedresults"]);
                 foldersToCreate.Add(mergedresults);

              publishresults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["publishresults"]);
              foldersToCreate.Add(publishresults);

              foreach (var folder in foldersToCreate)
                {
                    var dirinfo = Directory.CreateDirectory(folder);
     
                }

                // Files
                resultfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["results"]);
                startlistfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["startlist"]);
                sortedresultsfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["sortedresults"]);
                logo = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logo"]);
                preliminaryResults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["prel"]);
                logovoid = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logovoid"]);

                fakefile = Path.Combine(fakebox, "fakedresults.xlsx");

                if(!File.Exists(resultfile))
                {
                    showMessageBox("First time using folder " + workingDirectory + ". Copying base result file");
                    var ff = Path.Combine(Application.StartupPath, ConfigurationManager.AppSettings["results"]);
                    File.Copy(ff, resultfile);
                }

                if (!File.Exists(logo))
                {
                    var logoFile = Path.Combine(Application.StartupPath,"logos", ConfigurationManager.AppSettings["logo"]);
                    File.Copy(logoFile, logo);
                }
                if (!File.Exists(logovoid))
                {
                    var logoFile = Path.Combine(Application.StartupPath, "logos", ConfigurationManager.AppSettings["logovoid"]);
                    File.Copy(logoFile, logovoid);
                }

                if (!File.Exists(preliminaryResults))
                {
                    var preliminaryResultsFile = Path.Combine(Application.StartupPath, "logos", ConfigurationManager.AppSettings["prel"]);
                    File.Copy(preliminaryResultsFile, preliminaryResults);
                }





            }
            catch (Exception e)
            {
                showMessageBox(e.Message);
                Application.Exit();
            }
        }


        // read classes from startlist
        private List<Klass> readClasses()
        {
            if(!File.Exists(startlistfile))
            {
               UpdateMessageTextBox($"No startlist found, expecting " + startlistfile);
              showMessageBox("No startlist found, expecting " + startlistfile);
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
                   
                    //var f = dict[0, 2];

                    //var ws0 = (object[,])ws.Cells;


                    var count = row.Count();

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

                //classes = rows.Select(r => Klass.RowToClass(r)).ToList();
                //classes = truerows.Select(r => Klass.RowToClass(r)).ToList();


            }
          UpdateMessageTextBox($"Returning {classes.Count} classes");
          return classes;
        }

        // read deltagare from startlist
        private List<Deltagare> readVaulters()
        {
            if (!File.Exists(startlistfile))
            {
              UpdateMessageTextBox($"No startlist found, expecting " + startlistfile);
        showMessageBox("No startlist found, expecting " + startlistfile);
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

          UpdateMessageTextBox($"Returning {deltagare2.Count} vaulters" );
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


        private void SetFakedResult()
        {

        }

        /// <summary>
        /// Fake results by putting a value between 0 and 10 in the result cell of all sheets in Inbox
        /// </summary>
        private void doFake()
        {
            UpdateProgressBarHandler(0);
           
            UpdateProgressBarLabel("");

            DirectoryInfo dirinfo = new DirectoryInfo(fakebox);
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

            int fNumber = 0;
           

            UpdateProgressBarHandler(0);
            UpdateProgressBarLabel("");
            UpdateProgressBarMax(files.Count());


            // add a new worksheet to the empty workbook
            Dictionary<int, List<string>> data = new Dictionary<int, List<string>>();

            int rownumber = 0;

            foreach (var f1 in files)
            {
                rownumber++;
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

                        

                        var range = ws.NamedRange("result");
                        var refersTo = range.RefersTo;
                        var cellRange = ws.Range(refersTo);
                        var cells = cellRange.Cells();
                        cells.First().Value = rand;

                        
                        List<string> d = new List<string>();

                        refersTo = ws.NamedRange("klass").RefersTo;
                        var klassName = ws.Range(refersTo).Cells().First().Value.ToString();

                        refersTo = ws.NamedRange("bord").RefersTo;
                        var bord = ws.Range(refersTo).Cells().First().Value.ToString();

                        refersTo = ws.NamedRange("moment").RefersTo;
                        var moment = ws.Range(refersTo).Cells().First().Value.ToString();

                        refersTo = ws.NamedRange("id").RefersTo;
                        var id = ws.Range(refersTo).Cells().First().Value.ToString();

                        d.AddRange(new List<string> {klassName, f1.Name, id, moment, bord, rand.ToString(CultureInfo.InvariantCulture)});
                        data.Add(rownumber, d);
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
                    UpdateProgressBarHandler(rownumber);
                    UpdateProgressBarLabel("Faked " + f1.Name);
                }
            }


         

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
                            worksheet.Cells[row, i + 1].Value = float.Parse(kvp.Value[i]);
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



        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }


    //  public static void PrintToHtml(string className,string filename)
    //  {


    //  XLSXDocument docx = new XLSXDocument(@"C:\demoInput\demo.docx");

    //    //convert to html5 files
    //    docx.ConvertToVectorImages(ContextType.HTML, @"C:\htmlOutput\", "test", RelativeType.HTML);

    //    //convert to svg files
    //    //docx.ConvertToVectorImages(ContextType.SVG, @"C:\svgOutput\", "test", RelativeType.SVG);
    //}


    public Excel._Application printResultsExcelHandler(string className,string filename)
        {
            Excel.Application MyApp = null;
            Excel.Workbook MyBook = null;
            Excel.Workbooks workbooks = null;
            Excel.Worksheet MySheet = null;
            //Image prel = new Bitmap(preliminaryResults);

            try
            {
                MyApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false
                    
                    //DisplayAlerts = true
                };
                workbooks = MyApp.Workbooks;
                MyBook = workbooks.Open(sortedresultsfile,ReadOnly:true);
                MySheet = MyBook.Sheets[className];
                //MySheet.Activate();
                
                if (checkBox1.Checked)
                {
                    MySheet.PageSetup.RightHeaderPicture.Filename = preliminaryResults;
                }
                else
                {
                    MySheet.PageSetup.RightHeaderPicture.Filename = logovoid;
                }

                //MyApp.Visible = true;
                string fullpath = Path.Combine(printedresults, filename);
                MySheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, fullpath+".pdf");

                MyApp.DisplayAlerts = false;
                MyBook.Close();
                MyApp.DisplayAlerts = true;
                MyApp.Quit();
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

            return null;
        }


        // Export Results for class
        private void printResults(string className, string description)
        {
            
            this.UpdateMessageTextBox($"Save to PDF :  {className}");
            printResultsExcelHandler(className, description);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

           // PrintToHtml(className, description);
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
                printResults(cl.Name, cl.Name+"_"+cl.Description);
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
            UpdateMessageTextBox("Merging PDFs...");
            pdf.Merge(printedresults);
            UpdateMessageTextBox("Publishing results...");
            PDFtoHTML.GenerateHTML();
            UpdateMessageTextBox("Merge & Publish done...");
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
