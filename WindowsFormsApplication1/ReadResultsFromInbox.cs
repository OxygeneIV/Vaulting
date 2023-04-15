using OfficeOpenXml;
using System;
using System.Configuration;
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
            DirectoryInfo dirinfo = new DirectoryInfo(inboxFolder);
            var files = dirinfo.EnumerateFiles("*.xls*");
            var max = files.Count();
            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");

            if (files.Count() == 0)
            {
                UpdateMessageTextBox("No result files available");
                return;
            }

            UpdateMessageTextBox("Beginning import of results");
            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");

            var horseFileName = Form1.horseresultfile;
            FileInfo resultat = new FileInfo(resultfile);
            using (ExcelPackage results = new ExcelPackage(resultat))
            {
                try
                {
                    int counter = 0;
                    foreach (var f in files)
                    {
                        var toFile1 = Path.Combine(outboxFolder, f.Name);
                        if(File.Exists(toFile1))
                        {
                            // Overwrite
                            var msg = MessageBox.Show($@"File {Path.GetFileName(toFile1)} already exists in outbox!  Overwrite ?","",MessageBoxButtons.YesNo);
                            if (msg == DialogResult.Yes)
                            {
                                // continue using a backup
                                string date = DateTime.Now.ToString("yyyyMMddHHmmss");
                                string newfile = $"{toFile1}_{date}";
                                File.Move(toFile1,newfile);
                            }
                            else
                            {
                                UpdateMessageTextBox($"Ignoring file {f.Name}");
                                continue;
                            }
                           
                        }

                        // First copy the file before trying anything stupid.
                        var newPath = Path.Combine(backupFolder, f.Name);
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
                                    var id = ws.Cells["id"].Value.ToString();
                                    string refid = id;
                                    var refsplit = refid.Split('_');

                                    //var horsenumber = refsplit[3].Trim();


                                  // Horse analysis
                                  // SM & NM HorsePointStoring
                                  string horsename = null;
                                  try
                                  {
                                    var table = refsplit.Last().Trim();
                                    if (table.ToLower() == "a")
                                    {
                                      var datumcell = ws.Cells["datum"];
                                      var horsecell = datumcell.Offset(5, 0);
                                      horsename = horsecell.GetValue<string>().Trim();                                                    }
                                  }
                                  catch (Exception g)
                                  {
                                    UpdateMessageTextBox($"Failed to add horse point for {f.Name} , {g.Message}");
                                  }



                  // SM & NM
                  if (refid.Contains(".2")) // add results to 0 and 1
                  {
                    var klassMain = refsplit[3].Trim().Split('.').First();

                    var zero = refid.Replace(".2", "");
                    results.Workbook.Worksheets[klassMain].Cells[zero].Value = res;

                    var one = refid.Replace(".2", ".1");
                    results.Workbook.Worksheets[klassMain + ".1"].Cells[one].Value = res;

                    if (horsename != null)
                    {
                      File.AppendAllText(horseFileName, $"{refid};{horsename};{klassMain};{res}{Environment.NewLine}");
                      File.AppendAllText(horseFileName, $"{refid};{horsename};{klassMain + ".1"};{res}{Environment.NewLine}");
                    }

                  }
                  else
                  {

                    var klassMain = refsplit[2].Trim();

                    // Escamilo
                    /*
                    if (klassMain == "5" || klassMain == "6")
                    {
                        if (horsenumber == "")  // New Common Escamilo
                        {

                            // Update refid for class 5 and 6
                            var newRefsplit = refsplit;
                            newRefsplit[3]  = oldHorseNumber; // Old escamilo for class 5 and 6
                            refid = String.Join("_", newRefsplit).Trim();
                        }
                    }
                    */
                                       try { 
                                        results.Workbook.Worksheets[klassMain].Cells[refid].Value = res;
                                        }
                                        catch(Exception herr)
                                        {
                                             UpdateMessageTextBox("Failed to add result to ref "+klassMain + " " + refid+ " "+ f.Name);
                                        }
                                        if (horsename != null)
                                        {
                                          File.AppendAllText(horseFileName, $"{refid};{horsename};{klassMain};{res}{Environment.NewLine}");
                                        }
                                     }
                                }
                            }
                            var toFile = Path.Combine(outboxFolder, f.Name);
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
                    UpdateMessageTextBox("Completed import of results");
                    UpdateMessageTextBox("Completed import of results, saving...");
                    results.Save();
                    UpdateMessageTextBox("Save completed, wait for calculation");
                }
            }

            UpdateMessageTextBox("Import of results, calculating points...");
       
              bool docalc = Convert.ToBoolean(ConfigurationManager.AppSettings["resultcalchelper"]);
              if (!docalc)
              {
                UpdateMessageTextBox("Import of results, calculation done...sorting...");
                return;
              }
    
                var MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = true;
                var workbooks = MyApp.Workbooks;
                var MyBook = workbooks.Open(resultfile);
                MyApp.CalculateFull();
                MyBook.Close(true);
                workbooks.Close();
                MyApp.Quit();
                UpdateMessageTextBox("Import of results, calculation done...wait for sorting...");
                Marshal.ReleaseComObject(MyBook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(MyApp);
                MyBook = null;
                workbooks = null;
                MyApp = null;



        }
    }
}
