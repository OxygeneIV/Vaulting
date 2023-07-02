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
using SelectPdf;
using FluentFTP;
using System.Net;

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
    public static string htmlNoResultsFolder;

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

        if (competitionstarted)
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

        if (!Directory.Exists(logosFolder))
          logosFolder = Path.Combine(Application.StartupPath, "logos");

        //foldersToCreate.Add(logosfolder);

        printedresultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["printedresults"]);
        foldersToCreate.Add(printedresultsFolder);

        htmlResultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["htmlresultsfolder"]);
        foldersToCreate.Add(htmlResultsFolder);

        htmlNoResultsFolder = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["htmlNoResultsfolder"]);
        foldersToCreate.Add(htmlNoResultsFolder);

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
        omvandfile = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["omvandstartordning"]);
        //ridsportlogo = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logo"]);
        //preliminaryResults = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["prel"]);
        //logovoid = Path.Combine(workingDirectory, ConfigurationManager.AppSettings["logovoid"]);

        String cssfile = Path.Combine(Environment.CurrentDirectory, "html/stylesheet.css");
        File.Copy(cssfile, Path.Combine(cssFolder, "stylesm.css"), true);

        var files = Directory.EnumerateFiles(Path.Combine(Environment.CurrentDirectory, "html/img"));
        foreach (String f in files)
        {

          File.Copy(f, htmlResultsFolder + "/" + Path.GetFileName(f), true);
        }



        fakefile = Path.Combine(fakeboxFolder, "fakedresults.xlsx");

        if (!File.Exists(resultfile))
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
      if (!File.Exists(startlistfile))
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
          object[,] dict;
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
          bool isfloat = float.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out newSmclass);
          if (allEmpty || !(text.IsNumeric() || isfloat)) continue; // skip this row
          rows.Add(row);
          truerows.Add(trueRow);
          cellvals.Add(tmplist);
        }
        classes = cellvals.Select(r => Klass.RowToClass(r)).ToList();

        // Remove SM & NM
        classes.RemoveAll(c => c.Name.Contains("."));
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

          if (allEmpty || !(text.IsNumeric() || isfloat)) continue; // skip this row
          deltagarlista.Add(row);
        }
        deltagare = deltagarlista.Select(r => Deltagare.RowToClass(r)).ToList();


      }

      List<Deltagare> deltagare2 = new List<Deltagare>();

      // Double .2-people SM & NM

      foreach (var d in deltagare)
      {
        
        if (d.Klass.Contains("."))
        {
          List<String> theIds = d.Id.Split(',').ToList();
          List<String> theClasses = d.Klass.Split('.').ToList();

          if (theIds.Count != theClasses.Count)
            throw new Exception("Wrong sizes for klass and id " + d.Id);

          var d1 = d.Duplicate();
          d1.Klass = theClasses[0];
          d1.Id = theIds[0];
          deltagare2.Add(d1);

          var d2 = d.Duplicate();
          d2.Klass = theClasses[1];
          d2.Id = theIds[1];
          deltagare2.Add(d2);
        }
        else
        {
          deltagare2.Add(d.Duplicate());
        }
      }
      var allIds = deltagare2.Select(d => d.Id);
      var distinctIds = deltagare2.Select(d => d.Id).Distinct().Count();
      var duplicates = deltagare2.Count - distinctIds;



      UpdateMessageTextBox($"Found {deltagare2.Count} vaulters, {duplicates} duplicate IDs");
      //foreach (var d in deltagare2)
      //{
      //  UpdateMessageTextBox($"{allIds}");
      //}
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

      if (file.Exists)
      {
        file.Delete();
      }

      UpdateProgressBarHandler(0);
      UpdateProgressBarLabel("");
      UpdateProgressBarMax(files.Count());

      UpdateProgressBarHandler(0);
      UpdateProgressBarLabel("");
      UpdateProgressBarMax(files.Count());



      ExcelRange toRange;
      ExcelRange fromRange;
      int rownumber = 0;
      //Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 1 }, file1 =>
      //  {
      //      Interlocked.Increment(ref rownumber);
      //      using (var results = new ExcelPackage(file1))
      //      {
      //          try
      //          {

      //              var rand = Math.Round(new Random().NextDouble() * 10, 3);
      //              var sheet = results.Workbook.Worksheets.Single(w => w.Hidden.Equals(eWorkSheetHidden.Visible));
      //              fromRange = sheet.Cells["result"];
      //              fromRange.First().Value = rand;
      //              //fromRange.Value = rand;
      //                //fromRange.SetCellValue(0, 0, rand);
      //                //var adr = fromRange.Address;
      //                //sheet.SetValue(adr, rand);
      //            }
      //          catch (Exception e)
      //          {
      //              var s = e.Message;
      //              UpdateMessageTextBox($"Exception : {file1.Name} " + s);
      //          }
      //          finally
      //          {
      //              results.Save();
      //              UpdateProgressBarHandler(rownumber);
      //              UpdateProgressBarLabel("Faked " + rownumber + " " + file1.Name);
      //          }
      //      }
      //  });

      //return;

      ConcurrentDictionary<int, List<string>> data = new ConcurrentDictionary<int, List<string>>();



      Boolean woody = bool.Parse(ConfigurationManager.AppSettings["woody"]);

      rownumber = 0;
      foreach (var f1 in files)
      {
        //Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 50 }, f1 =>
        //{

        //Interlocked.Increment(ref rownumber);
        rownumber++;

        //lock (lockObject)
        //{
        UpdateProgressBarLabel("Faking # " + rownumber + " " + f1.Name);
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
            //if (f1.FullName.Contains("_A_") && woody)
            //{
            //  rand = 6.5;
            //}


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
              worksheet.Cells[row, i + 1].Value = float.Parse(kvp.Value[i], CultureInfo.InvariantCulture);
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

    public void UpdateMessageTextBoxWarn(string text)
    {
      if (this.textBox1.InvokeRequired)
        this.BeginInvoke(new UpdateMessageTextBoxCallback(this.UpdateMessageTextBoxWarn), new object[] { text });
      else
      {
        // change your text
        this.textBox1.ForeColor = System.Drawing.Color.Red;
        this.textBox1.AppendText(text + System.Environment.NewLine);// (char)13);
        this.textBox1.ForeColor = System.Drawing.Color.Black;

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
      if (!File.Exists(sortedresultsfile))
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


    public Excel._Application printResultsExcelHandler(string className, string filename)
    {
      Excel.Application MyApp = null;
      Excel.Workbook MyBook = null;
      Excel.Workbooks workbooks = null;
      Excel.Worksheet MySheet = null;
      bool preliminiaryResults = checkBox1.Checked;
      string fullpath = Path.Combine(printedresultsFolder, filename);
      fullpath = fullpath.Replace("*", "_star_");
      
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

            judgetext = judgetext + "   " + submom.Table.Name + " : " + judgeName;
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

        MySheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFullPath);

        MyApp.DisplayAlerts = false;
        MyBook.Close();
        MyApp.DisplayAlerts = true;
        MyApp.Quit();

        //createHtml(className);

      }
      catch (Exception e)
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

        PdfSharp.Pdf.PdfDocument document = PdfReader.Open(pdfFullPath, PdfDocumentOpenMode.Modify);

        for (int i = 0; i < document.Pages.Count; ++i)
        {
          PdfSharp.Pdf.PdfPage page = document.Pages[i];

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
              gfx.DrawImage(xim, new Point(500, 10));
            }
          }

        }

        document.Options.CompressContentStreams = true;
        document.Options.NoCompression = false;
        document.Save(pdfFullPath);
      }
      catch (Exception logoException)
      {
        this.UpdateMessageTextBox($"Save to PDF failed for {className} : {logoException.Message}");
      }

      return null;
    }


    private static string MakeFileNameWebSafe(string filename)
    {
      return filename.Replace(",", "-").Replace(" ", "-").Replace("å", "a").Replace("ü", "y");
    }



    private void backgroundWorkerPublish_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
    {
      UpdateMessageTextBox($"Publish Results BackgroundWorker Start...");
      createIndex();
      //createIndexNoPublish();
      UpdateMessageTextBox("Publishing results");
      publish();
      UpdateMessageTextBox("Publishing results completed");
      UpdateMessageTextBox($"Publish Results BackgroundWorker End...");
    }

    private void backgroundWorkerPublish_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
    {
      if (e.Cancelled == true)
      {
        //   "Canceled!";
      }
      else if (e.Error != null)
      {
        showMessageBox(e.Error.Message);
      }
      else
      {
        showMessageBox("Publish Results completed");
      }
    }



    public void publish()
    {
      
      var folder = Form1.htmlResultsFolder;
      var files = Directory.GetFiles(folder).ToList();
      var folders = Directory.GetDirectories(folder).ToList();

      // Test FTP
      var FTPserver = ConfigurationManager.AppSettings["ftpserver"];
      var FTPuser = ConfigurationManager.AppSettings["ftpuser"];
      var FTPpwd = ConfigurationManager.AppSettings["ftppwd"];
      var remoteworkingfolder = ConfigurationManager.AppSettings["remoteworkingfolder"];
      var remotepdfurl = ConfigurationManager.AppSettings["remotepdfurl"];

      try
      {
        FtpClient client = new FtpClient(FTPserver) { Credentials = new NetworkCredential(FTPuser, FTPpwd) };
        client.Connect();
        client.SetWorkingDirectory(remoteworkingfolder);
        //UploadFiles(localPaths, remoteDir, existsMode, createRemoteDir, verifyOptions, errorHandling)
        //client.UploadFiles(files, remoteworkingfolder, FtpRemoteExists.Overwrite);
        client.UploadDirectory(htmlResultsFolder, remoteworkingfolder, FtpFolderSyncMode.Update, FtpRemoteExists.Overwrite);
        client.Disconnect();
      }catch(Exception e)
      {
       
        UpdateMessageTextBox($"FTP failed...{e.Message}");

      }
    }


    private void createPdfFromHtml(String htmlFile)
    {
      bool pdfgeneration = createPdfsCheckBox.Checked;

      if (!pdfgeneration)
      {
        this.UpdateMessageTextBox("PDF creation not requested...");
        return;
      }
      String fullPdfName = htmlFile + ".pdf";
      String shortFile = Path.GetFileName(htmlFile);
      String pdfName = shortFile + ".pdf";
      this.UpdateMessageTextBox($"Creating PDF '{pdfName}' from HTML...");



      // Ta bort gamla pdf'er
      if (File.Exists(fullPdfName))
      {
        try
        {
          this.UpdateMessageTextBox($"Removing old PDF: {fullPdfName}");
          File.Delete(fullPdfName);
        }catch(Exception d)
        {
          this.UpdateMessageTextBox($"Delete old pdf failed...{d.Message}");
          return;
        }
      }

        this.UpdateMessageTextBox("Creating PDF...");
      try
      {
        HtmlToPdf converter = new HtmlToPdf();
        converter.Options.AutoFitWidth = HtmlToPdfPageFitMode.ShrinkOnly;

        // convert the url to pdf
        SelectPdf.PdfDocument doc = converter.ConvertUrl(htmlFile);

        // save pdf document
        doc.Save(fullPdfName);

        // close pdf document
        doc.Close();

        this.UpdateMessageTextBox("Creating PDF from HTML completed...");
      }
      catch(Exception d)
      {
        this.UpdateMessageTextBox($"Creating PDF from HTML failed... {d.Message}");
      }
    }


      private void createIndex()
    {

      this.UpdateMessageTextBox("Creating Indexfile...");

      String indexfile = Path.Combine(htmlResultsFolder, "index.html");
      File.Delete(indexfile);

 

      String headfile = Path.Combine(Environment.CurrentDirectory, "html/HTML_head.html");
      String mallIndex = Path.Combine(Environment.CurrentDirectory, "html/mallIndex.html");

      String head = File.ReadAllText(headfile);
      String text = File.ReadAllText(mallIndex);


      text = text.Replace("{HEAD}", head);

      var folder = Form1.htmlResultsFolder;
      var htmlfiles = Directory.GetFiles(folder, "*.html").ToList();
      htmlfiles.Sort(new PDFtoHTML.Comparer());

      bool pdfgeneration = createPdfsCheckBox.Checked;
      pdfgeneration = false;
     
      if (pdfgeneration)
      {
        this.UpdateMessageTextBox("Creating PDFs...");

        // Ta bort gamla pdf'er
        string[] filePaths = Directory.GetFiles(htmlResultsFolder, "*.pdf");
        foreach (string filePath in filePaths)
          File.Delete(filePath);

        HtmlToPdf converter = new HtmlToPdf();
        converter.Options.AutoFitWidth = HtmlToPdfPageFitMode.ShrinkOnly;


        foreach (String htmlFile in htmlfiles)
        {
          String shortFile = Path.GetFileName(htmlFile);

          // convert the url to pdf
          SelectPdf.PdfDocument doc = converter.ConvertUrl(htmlFile);

          // save pdf document
          doc.Save(htmlFile + ".pdf");

          // close pdf document
          doc.Close();
        }
      }
      else
      {
        this.UpdateMessageTextBox("Skipping PDFs...");
      }

      String ulLista = "";
      //ulLista = ulLista+ @"<table class=""table table-sm"">";
      //ulLista = ulLista + "<tbody>";

      long ticks = DateTime.Now.Ticks;
      String tickString = ticks.ToString();

      foreach (String htmlFile in htmlfiles)
      {
        ulLista = ulLista + "<tr>" + Environment.NewLine;
        String f = Path.GetFileName(htmlFile);
        f = f + "?" + tickString;

        String f2 = Path.GetFileNameWithoutExtension(htmlFile);
        f2 = f2.Replace("_star_", "*");
        String klassnum = f2.Split(' ')[0].Trim();

        String lnkformat = @"<td class=""indexunderline""><a href=""" + f +
                           @""">" + f2 + @"</a></td>" + Environment.NewLine;


        ulLista = ulLista + lnkformat + Environment.NewLine; ;
        /*
        String lnkformat2 = @"<td class=""indexunderline""><a href=""" + f + ".pdf" +
        @""">" + "PDF" + @"</a></td>" + Environment.NewLine;

        ulLista = ulLista + lnkformat2 + Environment.NewLine; ;
        */
        ulLista = ulLista + "</tr>" + Environment.NewLine;

      }

      //ulLista = ulLista + "</table>" + Environment.NewLine;

      text = text.Replace("{BODY}", ulLista);

      File.WriteAllText(indexfile, text,System.Text.Encoding.Unicode);

      this.UpdateMessageTextBox($"Creating Indexfile and PDFs completed...");

    }

    private void createIndexNoPublish()
    {

      this.UpdateMessageTextBox($"Creating Indexfile and PDFs for No Publish...");

      String indexfile = Path.Combine(htmlNoResultsFolder, "index.html");
      File.Delete(indexfile);

      // Ta bort gamla pdf'er
      string[] filePaths = Directory.GetFiles(htmlNoResultsFolder, "*.pdf");
      foreach (string filePath in filePaths)
        File.Delete(filePath);

      String headfile = Path.Combine(Environment.CurrentDirectory, "html/HTML_head.html");
      String mallIndex = Path.Combine(Environment.CurrentDirectory, "html/mallIndex.html");


      String head = File.ReadAllText(headfile);
      String text = File.ReadAllText(mallIndex);

      text = text.Replace("{HEAD}", head);

      var folder = Form1.htmlNoResultsFolder;
      var htmlfiles = Directory.GetFiles(folder, "*.html").ToList();
      htmlfiles.Sort(new PDFtoHTML.Comparer());

      bool pdfgeneration = createPdfsCheckBox.Checked;
      pdfgeneration = false;

      if (pdfgeneration)
      {
        this.UpdateMessageTextBox("Creating PDFs...");
        HtmlToPdf converter = new HtmlToPdf();
        converter.Options.AutoFitWidth = HtmlToPdfPageFitMode.ShrinkOnly;


        foreach (String htmlFile in htmlfiles)
        {
          String shortFile = Path.GetFileName(htmlFile);

          // convert the url to pdf
          SelectPdf.PdfDocument doc = converter.ConvertUrl(htmlFile);

          // save pdf document
          doc.Save(htmlFile + ".pdf");

          // close pdf document
          doc.Close();
        }
      }
      else
      {
        this.UpdateMessageTextBox("Skipping PDFs...");
      }

      String ulLista = "";
      //ulLista = ulLista+ @"<table class=""table table-sm"">";
      //ulLista = ulLista + "<tbody>";

      foreach (String htmlFile in htmlfiles)
      {
        ulLista = ulLista + "<tr>" + Environment.NewLine;
        String f = Path.GetFileName(htmlFile);
        String f2 = Path.GetFileNameWithoutExtension(htmlFile);
        f2 = f2.Replace("_star_", "*");
        String klassnum = f2.Split(' ')[0].Trim();

        String lnkformat = @"<td class=""indexunderline""><a href=""" + f +
                           @""">" + f2 + @"</a></td>" + Environment.NewLine;


        ulLista = ulLista + lnkformat + Environment.NewLine; ;
        /*
        String lnkformat2 = @"<td class=""indexunderline""><a href=""" + f + ".pdf" +
        @""">" + "PDF" + @"</a></td>" + Environment.NewLine;

        ulLista = ulLista + lnkformat2 + Environment.NewLine; ;
        */
        ulLista = ulLista + "</tr>" + Environment.NewLine;

      }

      //ulLista = ulLista + "</table>" + Environment.NewLine;

      text = text.Replace("{BODY}", ulLista);

      File.WriteAllText(indexfile, text);

      this.UpdateMessageTextBox($"Creating Indexfile and PDFs completed for No Publish...");

    }


    private String createHtml(String className)
    {
      String htmlFilePath = null;
      //  var deltagare = readVaulters();
      var classes = readClasses();
      //  var max = deltagare.Count();

      String nopublishString = ConfigurationManager.AppSettings["nopublish"];
      List<String> nopublishList = nopublishString.Split(',').ToList();

      //UpdateProgressBarHandler(0);
      //UpdateProgressBarMax(deltagare.Count);
      //UpdateProgressBarLabel("");

      var resultat = new FileInfo(sortedresultsfile);

      //FileInfo resultat = new FileInfo(resultfile);

      // keep track of first vaulter / class so we know if we shall copy range or not 
      List<string> set = new List<string>();
      Dictionary<string, int> vaulterInClassCounter = new Dictionary<string, int>();


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
        headfile = Path.Combine(Environment.CurrentDirectory, "html/HTML_head.html");

        String competition = Path.Combine(Environment.CurrentDirectory, "html/HTML_topCompetition.html");

        //cssfile = Path.Combine(Environment.CurrentDirectory, "html/stylesheet.css");

        if (klass.ResultTemplate.Equals("GK2"))
        {
          file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare2moment.html");
          file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header2domare2moment.html");
          _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat2domare2moment.html");
        }
        else if (klass.ResultTemplate.Equals("GKM3"))
        {
          file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare2moment.html");
          file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3_GKM3_domare2moment.html");
          _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat3_GKM3_domare2moment.html");
        }
        else if (klass.ResultTemplate.Equals("ResultTemplate"))
        {
          if (moments == 2)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top4domare2moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4domare2moment.html");
          }
          else if (moments == 3)
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
        else if (klass.ResultTemplate.Equals("GK3"))
        {
          if (moments == 2)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare2moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4_GK3_domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4_GK3_domare2moment.html");
          }
          else if (moments == 3)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare3moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3domare3moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat3domare3moment.html");
          }
          else if (moments == 4)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare4moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3domare4moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat3domare4moment.html");
          }
        }
        else if (klass.ResultTemplate.Equals("TRAGK2"))
        {
          if (moments == 2)  //HTML_resultat_tra_2-3domare1moment
          {
            //file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare2moment.html");
            //file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4_GK3_domare2moment.html");
            //_file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4_GK3_domare2moment.html");
            //HTML_resultat3_GKM3_domare2moment
            //HTML_header3_GKM3_domare2moment.html

            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare2moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3_GKM3_domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat_tra_2-3domare2moment.html");
          }
          else if (moments == 1) //Kyr
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare2moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3_GKM3_domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat_tra_2-3domare2moment.html");
          }
          else if (moments == 3)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare3moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3_GKM3_domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat_tra_2-3domare3moment.html");
          }
          else if (moments == 4)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare4moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3domare4moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat3domare4moment.html");
          }
        }
        else if (klass.ResultTemplate.Equals("TRAK1"))
        {
          if (moments == 1) //Kyr
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top2domare1moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3_GKM3_domare1moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat_tra_2-3domare1moment.html");
          }
        }
        else if (klass.ResultTemplate.Equals("GTK3"))
        {
          if (moments == 2)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare2moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4_GK3_domare2moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4_GK3_domare2moment.html");
          }
          else if (moments == 3)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare3moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header4_GTK3_domare3moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat4_GTK3_domare3moment.html");
          }
          else if (moments == 4)
          {
            file = Path.Combine(Environment.CurrentDirectory, "html/HTML_top3domare4moment.html");
            file2 = Path.Combine(Environment.CurrentDirectory, "html/HTML_header3domare4moment.html");
            _file3 = Path.Combine(Environment.CurrentDirectory, "html/HTML_resultat3domare4moment.html");
          }


        }

        else
        {
          throw new Exception(" no matching templates found for klass " + klass);
        }

        head = File.ReadAllText(headfile);
        text = File.ReadAllText(file);
        resultatheadertext = File.ReadAllText(file2);
        _text3 = File.ReadAllText(_file3);
        _text4 = File.ReadAllText(_file4);
        String textCompetition = File.ReadAllText(competition);



        bool preliminiaryResults = checkBox1.Checked;
        resultatheadertext = preliminiaryResults ? resultatheadertext.Replace("{HIDDEN}", "") : resultatheadertext.Replace("{HIDDEN}", "hidden");

        var sheet = results.Workbook.Worksheets[klass.Name];



        text = text.Replace("{KLASS}", "Klass " + klass.Name + " - " + klass.Description);
        int counter = 0;
        foreach (Moment moment in klass.Moments)
        {
          counter++;
          text = text.Replace("{MOMENT_" + counter + "}", moment.Name);
          resultatheadertext = resultatheadertext.Replace("{MOMENT_" + counter + "}", moment.Name);



          foreach (SubMoment submoment in moment.SubMoments)
          {
            String table = submoment.Table.Name;
            String judgename = submoment.Table.judge.Fullname;
            text = text.Replace("{MOMENT_" + counter + "_DOMARE_" + table + "}", judgename);
            resultatheadertext = resultatheadertext.Replace("{MOMENT_" + counter + "_" + table + "}", submoment.Name);
          }
        }

        //File.WriteAllText("test.html", text);
        //File.WriteAllText("test2.html", resultatheadertext);

        int rowbase = 7;
        int endrow = sheet.Dimension.End.Row;

        String textrows = "";

        String noresults = ConfigurationManager.AppSettings["noresults"];
        List<String> noresultsList = noresults.Split(',').ToList();
        Boolean noresultsInClass = noresultsList.Contains(klassnamn);

        int currentRowInTable = 0;
        int numberOfVaulters = (endrow - rowbase + 1 ) / 4;

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

          if(noresultsInClass && (placering.Trim() != "1"))
          {
            if(numberOfVaulters < 5)
            {
              if (currentRowInTable > 1)
                placering = "2";
            }
            else
            {
              if (currentRowInTable > 3)
                placering = "4";
            }
          }


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


            if (klass.ResultTemplate.Equals("GK3") || klass.ResultTemplate.Equals("GTK3"))
            {
              SubMoment s = new SubMoment();
              Table t = new Table();
              t.Name = "D";
              s.Table = t;
              moment.SubMoments.Add(s);
            }


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
              else if (point == null)
              {
                text3 = text3.Replace(keycell, "emptyExtra");
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

       // File.WriteAllText("test3.html", textrows);


        // Skapa fil
        _text4 = _text4.Replace("{HEAD}", head);
        _text4 = _text4.Replace("{COMPETITION}", textCompetition);
        _text4 = _text4.Replace("{TOP}", text);
        _text4 = _text4.Replace("{HEADER}", resultatheadertext);
        _text4 = _text4.Replace("{DATA}", textrows);
        String desc = klass.Description.Replace("*", "_star_");

        

        if (nopublishList.Contains(klass.Name.Trim()))
        {
          htmlFilePath = Path.Combine(htmlNoResultsFolder, klass.Name + " - " + desc + ".html");
          File.WriteAllText(Path.Combine(htmlNoResultsFolder, klass.Name + " - " + desc + ".html"), _text4);
        }
        else 
        {
          htmlFilePath = Path.Combine(htmlResultsFolder, klass.Name + " - " + desc + ".html");
          File.WriteAllText(Path.Combine(htmlResultsFolder, klass.Name + " - " + desc + ".html"), _text4);
        }

        int h = 5;


      }

      UpdateProgressBarLabel("All vaulters added to result file");

   
      return htmlFilePath;


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
      String htmlPath = null;
      try
      {
        UpdateMessageTextBox($"Saving class '{className}' to HTML");
        htmlPath = createHtml(className);
        UpdateMessageTextBox($"Saving class '{className}' to HTML done...");
      }
      catch (Exception ee)
      {
        UpdateMessageTextBox($"Saving class {className} to HTML failed...");
        UpdateMessageTextBox(ee.Message);
      }

      if (htmlPath != null)
      {
        createPdfFromHtml(htmlPath);
      }
      else
      {
        UpdateMessageTextBox($"Could not create PDF from HTML...");
      }
      /*
      bool pdfgeneration = createPdfsCheckBox.Checked;

      if (pdfgeneration)
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
      }
      */


        GC.Collect();
      GC.WaitForPendingFinalizers();
      GC.Collect();
      GC.WaitForPendingFinalizers();
    }

    // Export Results for selected class
    private void button2_Click_1(object sender, EventArgs e)
    {
      ClassSelect sel = comboBox1.SelectedItem as ClassSelect;
      string value = sel.Value;
      string text = sel.Text;
      printResults(value, text);

    }

    private void backgroundWorkerPrintResults_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
    {
      UpdateMessageTextBox($"Print Results BackgroundWorker Start...");
      this.doPrintResults();
      UpdateMessageTextBox($"Print Results BackgroundWorker End...");
    }

    private void backgroundWorkerPrintResults_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
    {
      if (e.Cancelled == true)
      {
        //   "Canceled!";
      }
      else if (e.Error != null)
      {
        showMessageBox(e.Error.Message);
      }
      else
      {
        showMessageBox("Print Results completed");
      }
    }

    private void doPrintResults()
    {
      var allClasses = readClasses();
      foreach (var cl in allClasses)
      {
        printResults(cl.Name, cl.Name + " " + cl.Description);
      }
    }

    // Export Results for all classes
    private void button5_Click(object sender, EventArgs e)
    {

      backgroundWorkerPrintResults.RunWorkerAsync();
      bool hasAllThreadsFinished = false;
      while (!hasAllThreadsFinished)
      {
        hasAllThreadsFinished = backgroundWorkerPrintResults.IsBusy == false;
        Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                //from the Progress event of the background worker.
        System.Threading.Thread.Sleep(100);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                //re-checking.
      }

      //var allClasses = readClasses();
      //foreach (var cl in allClasses)
      //{
      //  printResults(cl.Name, cl.Name + " " + cl.Description);
      //}



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

    //Publicera
    private void button1_Click(object sender, EventArgs e)
    {

      createIndex();
      createIndexNoPublish();
      UpdateMessageTextBox("Publishing results");
      publish();
      UpdateMessageTextBox("Publishing results completed");

      //try
      //{
      //  UpdateMessageTextBox("Merging PDFs...");
      //  pdf.Merge(printedresultsFolder);
      //  UpdateMessageTextBox("Merging PDFs done...");
      //}
      //catch (Exception ee)
      //{
      //  UpdateMessageTextBox("Failed to Merge PDFs ...");
      //  UpdateMessageTextBox(ee.Message);
      //}

      //try
      //{
      //  UpdateMessageTextBox("Publishing results...");
      //  PDFtoHTML.GenerateHTML();
      //  UpdateMessageTextBox("Publish done...");
      //}
      //catch(Exception ee)
      //{
      //  UpdateMessageTextBox("Failed to Publish ...");
      //  UpdateMessageTextBox(ee.Message);
      //}
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

      public bool IsSM => Int32.Parse(Klass)<20;
      public bool IsNM => Int32.Parse(Klass) >= 20;
      public bool IsSMNM => Id.Contains("A");

      public static HPclass Create(string hpline)
      {
        var data = hpline.Split(';');
        var hp = new HPclass
        {
          Id = data[0].Trim(),
          Name = data[1].Trim(),
          Klass = data[2].Trim(),
          point = float.Parse(data[3].Trim().Replace(",", "."), CultureInfo.InvariantCulture)
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
      var horsepointsCalculated = Path.Combine(Form1.horseResultsFolder, "CalculatedHorsePoints.xlsx");
      var horsepointsCalculatedTemplate = Path.Combine(Application.StartupPath, "CalculatedHorsePoints_template.xlsx");


      var allHPs = File.ReadAllLines(horsepoints).Distinct().Select(HPclass.Create).ToList();
      var removedhorsepoints = allHPs.RemoveAll(hp => !horsepointclasses.Contains(hp.Klass));
      UpdateMessageTextBox($"Removed {removedhorsepoints} from calculation");


      //var allPointsInd  = allHPs.Where(hp => !teamclasses.Contains(hp.Klass));
      //var allPointsTeam = allHPs.Where(hp => teamclasses.Contains(hp.Klass));



      File.Delete(horsepointsCalculated);
      File.Copy(horsepointsCalculatedTemplate, horsepointsCalculated, true);

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
                                SMNMMax = allSMNMPoints.Where(hp => hp.Name == AllHorsePoints.Key).Max(s => s.point),
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
      catch (Exception e)
      {
        UpdateMessageTextBox($"Failed to analyze horse points : {e.Message}");
      }

    }




    //public void CalculateHorsePoints()
    //{
    //  var teamclasses = ConfigurationManager.AppSettings["teamclasses"].Split(',').Select(s => s.Trim());

    //  UpdateMessageTextBox($"Starting horse point calculation");
    //  var resultfile = Form1.sortedresultsfile;
    //  var horsefile = horseresultfile;

    //  if (!File.Exists(resultfile))
    //  {
    //    UpdateMessageTextBox($"{resultfile} not found, aborting horse point calculation");
    //    return;
    //  }

    //  FileInfo resultat = new FileInfo(resultfile);
    //  FileInfo horsefileInfo = new FileInfo(horsefile);

    //  List<string> horses = new List<string>();

    //  List<Horse> definedHorses = new List<Horse>();
    //  List<Horse> definedHorsesTeam = new List<Horse>();
    //  List<Horse> definedHorsesInd = new List<Horse>();

    //  var classes = readClasses();

    //  using (ExcelPackage results = new ExcelPackage(resultat))
    //  {
    //    try
    //    {
    //      foreach (var cl in classes)
    //      {
    //        UpdateMessageTextBox($"Getting horse points from class {cl.Name} - {cl.Description}");
    //        int startRow = 7;
    //        ExcelWorksheet ws = results.Workbook.Worksheets[cl.Name];
    //        var maxrow = ws.Dimension.End.Row;

    //        int ekipages = (maxrow - startRow + 1) / 4;

    //        for (int ekipage = 0; ekipage < ekipages; ekipage++)
    //        {
    //          var currentStartRow = startRow + (ekipage * 4);
    //          var horsename = ws.Cells[currentStartRow + 2, 6].Value.ToString();
    //          horses.Add(horsename);

    //          if (!definedHorses.Any(h => h.Name == horsename))
    //          {
    //            Horse h1 = new Horse();
    //            h1.Name = horsename;
    //            definedHorses.Add(h1);
    //          }


    //          // TEAM
    //          if (teamclasses.Contains(ws.Name))
    //          {
    //            if (!definedHorsesTeam.Any(h => h.Name == horsename))
    //            {
    //              Horse h1 = new Horse();
    //              h1.Name = horsename;
    //              definedHorsesTeam.Add(h1);
    //            }
    //          }
    //          // IND
    //          else
    //          {
    //            if (!definedHorsesInd.Any(h => h.Name == horsename))
    //            {
    //              Horse h1 = new Horse();
    //              h1.Name = horsename;
    //              definedHorsesInd.Add(h1);
    //            }
    //          }

    //          var curhorse = definedHorses.Single(h => h.Name == horsename);

    //          for (int arow = 0; arow < 4; arow++)
    //          {
    //            var momenttext = ws.Cells[currentStartRow + arow, 7].Value.ToString();
    //            if (momenttext.Length > 1) // we may have points
    //            {
    //              var point = ws.Cells[currentStartRow + arow, 8].GetValue<float>();
    //              if (point > 0)
    //              {
    //                curhorse.Points.Add(point);

    //                // TEAM
    //                if (teamclasses.Contains(ws.Name))
    //                {
    //                  var curhorse1 = definedHorsesTeam.Single(h => h.Name == horsename);
    //                  curhorse1.Points.Add(point);

    //                }
    //                // IND
    //                else
    //                {
    //                  var curhorse2 = definedHorsesInd.Single(h => h.Name == horsename);
    //                  curhorse2.Points.Add(point);
    //                }
    //              }
    //            }
    //          }
    //        }
    //      }

    //    }
    //    catch (Exception ex)
    //    {
    //      var str = ex.Message;
    //      UpdateMessageTextBox(str);
    //    }
    //    finally
    //    {

    //    }
    //  }

    //  UpdateMessageTextBox($"Getting horse points from all classes done");
    //  var all = horses.Distinct().ToList();
    //  all.RemoveAll(s => s == "A4");
    //  definedHorses.RemoveAll(h => h.Name == "A4");
    //  definedHorsesTeam.RemoveAll(h => h.Name == "A4");
    //  definedHorsesInd.RemoveAll(h => h.Name == "A4");

    //  definedHorses.Sort();
    //  definedHorsesTeam.Sort();
    //  definedHorsesInd.Sort();

    //  File.Delete(horsefileInfo.FullName);

    //  using (ExcelPackage results = new ExcelPackage(horsefileInfo))
    //  {
    //    try
    //    {
    //      var sheet = results.Workbook.Worksheets.Add("Horse points team+ind");
    //      var sheet2 = results.Workbook.Worksheets.Add("Horse points team");
    //      var sheet3 = results.Workbook.Worksheets.Add("Horse points ind");
    //      sheet.Cells.Style.Numberformat.Format = @"0.000";
    //      sheet.Cells[1, 1].Value = "Häst";
    //      sheet.Cells[1, 3].Value = "Högsta enskilda poäng";
    //      sheet.Cells[1, 2].Value = "Medelpoäng";
    //      sheet.Cells[1, 4].Value = "Samtliga poäng";

    //      sheet2.Cells.Style.Numberformat.Format = @"0.000";
    //      sheet2.Cells[1, 1].Value = "Häst";
    //      sheet2.Cells[1, 3].Value = "Högsta enskilda poäng";
    //      sheet2.Cells[1, 2].Value = "Medelpoäng";
    //      sheet2.Cells[1, 4].Value = "Samtliga poäng";

    //      sheet3.Cells.Style.Numberformat.Format = @"0.000";
    //      sheet3.Cells[1, 1].Value = "Häst";
    //      sheet3.Cells[1, 3].Value = "Högsta enskilda poäng";
    //      sheet3.Cells[1, 2].Value = "Medelpoäng";
    //      sheet3.Cells[1, 4].Value = "Samtliga poäng";

    //      int row = 1;

    //      foreach (Horse h in definedHorses)
    //      {
    //        row = row + 1;
    //        sheet.Cells[row, 1].Value = h.Name;
    //        sheet.Cells[row, 3].Value = h.Max;
    //        sheet.Cells[row, 2].Value = h.Average;
    //        for (int i = 0; i < h.Points.Count; i++)
    //        {
    //          sheet.Cells[row, 4 + i].Value = h.Points[i];
    //        }

    //      }
    //      sheet.Cells.AutoFitColumns();

    //      row = 1;
    //      foreach (Horse h in definedHorsesTeam)
    //      {
    //        row = row + 1;
    //        sheet2.Cells[row, 1].Value = h.Name;
    //        sheet2.Cells[row, 3].Value = h.Max;
    //        sheet2.Cells[row, 2].Value = h.Average;
    //        for (int i = 0; i < h.Points.Count; i++)
    //        {
    //          sheet2.Cells[row, 4 + i].Value = h.Points[i];
    //        }

    //      }
    //      sheet2.Cells.AutoFitColumns();

    //      row = 1;
    //      foreach (Horse h in definedHorsesInd)
    //      {
    //        row = row + 1;
    //        sheet3.Cells[row, 1].Value = h.Name;
    //        sheet3.Cells[row, 3].Value = h.Max;
    //        sheet3.Cells[row, 2].Value = h.Average;
    //        for (int i = 0; i < h.Points.Count; i++)
    //        {
    //          sheet3.Cells[row, 4 + i].Value = h.Points[i];
    //        }

    //      }
    //      sheet3.Cells.AutoFitColumns();
    //      UpdateMessageTextBox($"{horsefile} created ! ");
    //    }
    //    catch (Exception ex)
    //    {
    //      UpdateMessageTextBox($"Horse point Error! ");
    //      UpdateMessageTextBox(ex.Message);
    //    }
    //    finally
    //    {
    //      results.Save();
    //      UpdateMessageTextBox($"{horsefile} saves ! ");
    //    }
    //  }

    //}

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

    private void checkBox2_CheckedChanged(object sender, EventArgs e)
    {

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
