using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{
    partial class Form1
    {

        private void backgroundWorkerSortResults_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            this.SortResults();
        }

        private void backgroundWorkerSortResults_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
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
                showMessageBox("All results sorted in "+ Path.GetFileName(sortedresultsfile));
            }
        }


    /// <summary>
    /// Sort results. If argument sort only this class
    /// </summary>
    /// <param name="inklass"></param>
    private void SortResults(string inklass = null)
        {
         
            if (File.Exists(sortedresultsfile))
            {
                File.Delete(sortedresultsfile);
            }

            var classes = readClasses();
            var max = classes.Count();

            if (inklass != null)
            {
                classes = classes.Where(c => c.Name == inklass).ToList();
            }

            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");
            UpdateProgressBarLabel("Starting Sort!!");
            UpdateMessageTextBox("Starting Sort of results...");
            File.Copy(resultfile, sortedresultsfile);

            var MyApp = new Application();
            MyApp.Visible = false;
            var workbooks = MyApp.Workbooks;
            var MyBook = workbooks.Open(sortedresultsfile);
            
            int counter = 0;


            foreach (Klass klass in classes)
            {
                counter++;
                string className = klass.Name;
                var MySheet = MyBook.Sheets[className];
                
                MySheet.Activate();
              UpdateMessageTextBox($"Sorting {className}");
        var lastRow = MySheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
                Microsoft.Office.Interop.Excel.Range newRng = MySheet.Range[MySheet.Cells[7, 1], MySheet.Cells[lastRow, 15]];
                newRng.Sort(
                            newRng.Columns[1, Type.Missing], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                newRng.Columns[2, Type.Missing], Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                            Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
                                XlSortOrientation.xlSortColumns,
                                Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                                Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                UpdateProgressBarHandler(counter);
                UpdateProgressBarLabel("Sorted class ( " + counter + " / " + max + " ) " + klass.Name + " - " + klass.Description);
            }

            MyBook.Close(true);
            workbooks.Close();
            MyApp.Quit();

            Marshal.ReleaseComObject(MyBook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(MyApp);
            MyBook = null;
            workbooks = null;
            MyApp = null;

            UpdateProgressBarLabel("Sorting completed");
            UpdateMessageTextBox($"Sorting completed");
    }
    }
}
