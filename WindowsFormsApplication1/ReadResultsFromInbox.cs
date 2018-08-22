using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    partial class Form1
    {

        private void showMessageBox(string text)
        {
            if (showmessageboxes)
            {
                //this.BeginInvoke((Action)(() => MessageBox.Show(text)));
                MessageBox.Show(this, text);
            }
        }

        private void backgroundWorkerReadResultsFromInbox_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            this.ReadResultsFromInbox();
        }

        private void backgroundWorkerReadResultsFromInbox_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
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
                showMessageBox("Imported results from Inbox, sorting will start after closing this dialog...");
            }
        }

        // Populate results
        private void btnReadResultsFromInbox_Click(object sender, EventArgs e)
        {
            backgroundWorkerReadResultsFromInbox.RunWorkerAsync();
            bool hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerReadResultsFromInbox.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(100);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                       //re-checking.
            }

            backgroundWorkerSortResults.RunWorkerAsync();
            hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerSortResults.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(100);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                        //re-checking.
            }


        }

        private void doSort()
        {
            bool hasAllThreadsFinished = false;
            backgroundWorkerSortResults.RunWorkerAsync();
            hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerSortResults.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(100);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                        //re-checking.
            }

        }

        private void ReadResultsFromInbox()
        {
            DirectoryInfo dirinfo = new DirectoryInfo(inbox);
            var files = dirinfo.EnumerateFiles("*.xls*");
            var max = files.Count();
            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");

            if (files.Count() == 0)
            {
                showMessageBox("No result files available");
                return;
            }
            UpdateMessageTextBox("Beginning import of results");

            UpdateProgressBarHandler(0);
            UpdateProgressBarLabel("");
            UpdateProgressBarMax(files.Count());



            // If we have issues with excel versions, try by enabling this foreach loop and "int fNumber" definition
            // Opens and saves the file using local excel application

            //int fNumber = 0;
            //foreach (var f in files)
            //{
            //    // First copy the file before trying anything stupid.
            //    var newPath = Path.Combine(backup, f.Name);
            //    File.Copy(f.FullName, newPath, true);

            //    fNumber++;
            //    UpdateProgressBarLabel("Re-saving for compatibility issues : " + f.Name);
            //    var MyApp = new Microsoft.Office.Interop.Excel.Application();
            //    MyApp.Visible = false;
            //    var workbooks = MyApp.Workbooks;
            //    var MyBook = workbooks.Open(f.FullName);
            //    MyBook.Close(true);
            //    workbooks.Close();
            //    MyApp.Quit();

            //    Marshal.ReleaseComObject(MyBook);
            //    Marshal.ReleaseComObject(workbooks);
            //    Marshal.ReleaseComObject(MyApp);
            //    MyBook = null;
            //    workbooks = null;
            //    MyApp = null;
            //    UpdateProgressBarHandler(fNumber);
            //}

            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");

            FileInfo resultat = new FileInfo(resultfile);
            using (ExcelPackage results = new ExcelPackage(resultat))
            {
                //results.Workbook.CalcMode = ExcelCalcMode.Automatic;
                try
                {
                    int counter = 0;
                    foreach (var f in files)
                    {

                        //    // First copy the file before trying anything stupid.
                        var newPath = Path.Combine(backup, f.Name);
                        File.Copy(f.FullName, newPath, true);

                        counter++;
                        try
                        {
                            using (ExcelPackage p = new ExcelPackage(new FileInfo(f.FullName)))
                            {
                                ExcelWorksheet ws = p.Workbook.Worksheets.Single(s => s.Hidden == eWorkSheetHidden.Visible);
                                bool omd = ws.Name.ToLower().EndsWith(" omd");
                                if (omd)
                                {

                                }
                                else
                                {
                                    var res = ws.Cells["result"].GetValue<float>();
                                    var klassName = ws.Cells["klass"].Value.ToString();
                                    var bord = ws.Cells["bord"].Value.ToString();
                                    var moment = ws.Cells["moment"].Value.ToString();
                                    var id = ws.Cells["id"].Value.ToString();

                                    // ID
                                    //string refid = id + "_" + klassName + "_" + moment.Replace(' ', '_') + "_" + bord;
                                    string refid = id;

                                    // SM & NM
                                    if (refid.Contains(".2")) // add results to 0 and 1
                                    {
                                        // 0
                                        var refsplit = refid.Split('_');
                                        var klassMain = refsplit[3].Trim().Split('.').First();

                                        var zero = refid.Replace(".2", "");
                                        results.Workbook.Worksheets[klassMain].Cells[zero].Value = res;

                                        var one = refid.Replace(".2", ".1");
                                        results.Workbook.Worksheets[klassMain+".1"].Cells[one].Value = res;
                                    }
                                    else
                                    {
                                        var refsplit = refid.Split('_');
                                        var klassMain = refsplit[3].Trim();

                                        results.Workbook.Worksheets[klassMain].Cells[refid].Value = res;
                                    }
                                }
                            }
                            var toFile = Path.Combine(outbox, f.Name);
                            File.Move(f.FullName, toFile);
                        }
                        catch (Exception e)
                        {
                            UpdateMessageTextBox("Error reading result from " + f.FullName);
                            UpdateMessageTextBox("  -> " + e.Message);
                            UpdateMessageTextBox("  -> " + e.StackTrace);
                        }

                        UpdateProgressBarHandler(counter);
                        UpdateProgressBarLabel("Read result from file ( " + counter + " / " + max + " ) " + f.Name);
                    }
                 }
                 catch(Exception e)
                 {
                    UpdateMessageTextBox("Error occured during result import !...");
                    UpdateMessageTextBox(e.Message);
                }
                finally
                {
                    UpdateMessageTextBox("Completed import of results, calculate..");
                    //var calcOptions = new ExcelCalculationOption();                    
                    //results.Workbook.Calculate(new ExcelCalculationOption());
                    UpdateMessageTextBox("Completed import of results, saving...");
                    results.Save();
                    //var calcOptions = new ExcelCalculationOption();                    
                    //results.Workbook.Calculate();
                    UpdateMessageTextBox("Save completed");
                }
            }

            var MyApp = new Microsoft.Office.Interop.Excel.Application();
            MyApp.Visible = true;
            var workbooks = MyApp.Workbooks;
            var MyBook = workbooks.Open(resultfile);
            MyApp.CalculateFull();
            MyBook.Close(true);
            workbooks.Close();
            MyApp.Quit();

            Marshal.ReleaseComObject(MyBook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(MyApp);
            MyBook = null;
            workbooks = null;
            MyApp = null;



        }
    }
}
