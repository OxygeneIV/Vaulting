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

        //private void SortResultsClosedXML(string inklass = null)
        //{
        //    if (File.Exists(sortedresultsfile))
        //    {
        //        File.Delete(sortedresultsfile);
        //    }
        //    var classes = readClasses();

        //    File.Copy(resultfile, sortedresultsfile);

        //    using (ExcelPackage p = new ExcelPackage(new FileInfo(sortedresultsfile)))
        //    {
                //foreach (Klass klass in classes)
                //{
                //    string className = klass.Name;
                //    var classWorksheet = p.Workbook.Worksheets[className];
                //    var upper = classWorksheet.Dimension.End.Row;
                //    var rng = classWorksheet.Cells[7, 1, upper, 1];
                //    //rng.Calculate();

                //    foreach (var cell in rng)
                //    {
                //        var v = cell.Text;

                //        cell.Formula = string.Empty;
                //        cell.Value = v;
                //    }
                //    //foreach (var cell in rng.Where(cell => cell.Formula != null))
                //    //    cell.Value = cell.Value;
                //}
                //p.Workbook.Calculate();
                //p.Workbook.CalcMode = ExcelCalcMode.Manual;
                //p.Save();
            //}

            //using (XLWorkbook workbook = new XLWorkbook(sortedresultsfile))
            //{
            //    var max = classes.Count();

            //    if (inklass != null)
            //    {
            //        classes = classes.Where(c => c.Name == inklass).ToList();
            //    }

            //    int counter = 0;

            //    foreach (Klass klass in classes)
            //    {
            //        counter++;
            //        string className = klass.Name;

            //        using (IXLWorksheet ws1 = workbook.Worksheet(className))
            //        {
            //            var maxrow = ws1.LastRowUsed().RowNumber();
            //            var myRange = ws1.Range(7, 1, maxrow, 1);
            //            foreach(var cell in myRange.Cells())
            //            {
            //                var val = cell.ValueCached;
            //            }
            //        }
            //    }
            //    workbook.Save();
            //}


            //using (XLWorkbook workbook = new XLWorkbook(sortedresultsfile))
            //{
            //    //workbook.CalculationOnSave = false;
            //    //var classes = readClasses();
            //    var max = classes.Count();

            //    if (inklass != null)
            //    {
            //        classes = classes.Where(c => c.Name == inklass).ToList();
            //    }

            //    UpdateProgressBarHandler(0);
            //    UpdateProgressBarMax(max);
            //    UpdateProgressBarLabel("");

            //    int counter = 0;

            //    foreach (Klass klass in classes)
            //    {
            //        counter++;
            //        string className = klass.Name;

            //        using (IXLWorksheet ws1 = workbook.Worksheet(className))
            //        {
                      
            //            var maxrow = ws1.LastRowUsed().RowNumber();
            //            var myRange = ws1.Range(7, 1, maxrow, 15);
            //            myRange.Sort();
            //        }


            //        UpdateProgressBarHandler(counter);
            //        UpdateProgressBarLabel("Sorted class ( " + counter + " / " + max + " ) " + klass.Name + " - " + klass.Description);
            //    }

            //    workbook.Save();
            //}
        //}
    
    


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
            File.Copy(resultfile, sortedresultsfile);

            var MyApp = new Microsoft.Office.Interop.Excel.Application();
            MyApp.Visible = true;
            var workbooks = MyApp.Workbooks;
            var MyBook = workbooks.Open(sortedresultsfile);
            
            int counter = 0;


            foreach (Klass klass in classes)
            {
                counter++;
                string className = klass.Name;
                var MySheet = MyBook.Sheets[className];

                MySheet.Activate();
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
        }
    }
}
