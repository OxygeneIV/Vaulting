using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    partial class Form1
    {

        private void backgroundWorkerPopulateSheetsWithVaulters_DoWork(object sender, DoWorkEventArgs e)
        {
            this.PopulateResultSheetsWithVaulters();
        }
        private void backgroundWorkerPopulateSheetsWithVaulters_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
                showMessageBox("Populated all result sheets with vaulters from startlist");
            }
        }

        private void buttonPopulateSheetsWithVaulters_Click(object sender, EventArgs e)
        {
            backgroundWorkerPopulateSheetsWithVaulters.RunWorkerAsync();
            bool hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerPopulateSheetsWithVaulters.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(100);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                       //re-checking.
            }
        }


        private void AddVaulterToResultsFile()
        {

        }

        // Load vaulters from startlist into result lists in correct class sheets
        private void PopulateResultSheetsWithVaulters()
        {
            var deltagare = readVaulters();
            var classes = readClasses();
            var max = deltagare.Count();

            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(deltagare.Count);
            UpdateProgressBarLabel("");

            FileInfo resultat = new FileInfo(resultfile);

            // keep track of first vaulter / class so we know if we shall copy range or not 
            List<string> set = new List<string>();
            Dictionary<string, int> vaulterInClassCounter = new Dictionary<string, int>();
            foreach(Klass c in classes)
            {
                vaulterInClassCounter[c.Name] = 0;
            }
            

            ExcelRange toRange;
            ExcelRange fromRange;

            using (var results = new ExcelPackage(resultat))
            {
                int deltagarCounter = 0;
                foreach (Deltagare d in deltagare)
                {
                    deltagarCounter++;
                    var klass = d.Klass;



                    var sheet = results.Workbook.Worksheets[klass];
                    //sheet.HeaderFooter.differentFirst = true; //ML
                    //sheet.HeaderFooter.differentOddEven = false;
                    int row = sheet.Dimension.End.Row;

                    fromRange = sheet.Cells["ekipage"];


                    // We have more than 1 competitor in the class and need to copy the ekipage range
                    //if (set.Contains(klass))
                    if (vaulterInClassCounter[klass]>0)
                    {
                        toRange = sheet.Cells[row + 1, 1, row + 4, fromRange.End.Column];
                        fromRange.Copy(toRange);


                        //Set formatting
                        for (int i = 1; i < 5; i++)
                        {
                            var theclass = classes.Single(c => c.Name == klass);
                            var endcol = 11;

                            if (theclass.ResultTemplate.Trim().EndsWith("1"))
                            {
                                endcol = 8;
                                ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                                _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                                _cond1_2.StopIfTrue = true;

                                ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                                _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                            }

                            if (theclass.ResultTemplate.Trim().EndsWith("2"))
                            {
                                endcol = 9;
                                ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                                _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                                _cond1_2.StopIfTrue = true;

                                ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                                _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                            }

                            if (theclass.ResultTemplate.Trim().EndsWith("M3"))
                            {
                                endcol = 10;
                                ExcelAddress _formatRangeAddress_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond1_2 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress_2);
                                _cond1_2.Formula = $"COUNTBLANK($G{row + i})=1";
                                _cond1_2.StopIfTrue = true;

                                ExcelAddress _formatRangeAddress2_2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond2_2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2_2);
                                _cond2_2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                _cond2_2.Style.Fill.BackgroundColor.Index = 3;
                            }


                            if (theclass.ResultTemplate.Trim().EndsWith("K3"))
                            {
                                endcol = 10;

                                ExcelAddress _formatRangeAddress = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond1 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress);
                                _cond1.Formula = $"COUNTBLANK($G{row + i})=1";
                                _cond1.StopIfTrue = true;

                                ExcelAddress _formatRangeAddress2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2);
                                _cond2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                _cond2.Style.Fill.BackgroundColor.Index = 3;

                                endcol = 11;
                                ExcelAddress _formatRangeAddressB = new ExcelAddress(row + i, endcol, row + i, endcol);
                                var _cond1B = sheet.ConditionalFormatting.AddExpression(_formatRangeAddressB);
                                _cond1B.Formula = $"COUNTBLANK($G{row + i})=0";
                                _cond1B.StopIfTrue = false;
                                _cond1B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                var color = System.Drawing.ColorTranslator.FromHtml("#E2EBD5");
                                //_cond1B.Style.Fill.BackgroundColor.Index = -4142;
                                _cond1B.Style.Fill.BackgroundColor.Color = color;

                                //ExcelAddress _formatRangeAddress2B = new ExcelAddress(row + i, endcol, row + i, endcol);
                                //var _cond2B = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2B);
                                //_cond2B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                //_cond2B.Style.Fill.BackgroundColor.Index = 35;

                            }

                            if (theclass.ResultTemplate.Trim().EndsWith("ResultTemplate"))
                            {
                                endcol = 11;

                                ExcelAddress _formatRangeAddress = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond1 = sheet.ConditionalFormatting.AddExpression(_formatRangeAddress);
                                _cond1.Formula = $"COUNTBLANK($G{row + i})=1";
                                _cond1.StopIfTrue = true;

                                ExcelAddress _formatRangeAddress2 = new ExcelAddress(row + i, 8, row + i, endcol);
                                var _cond2 = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2);
                                _cond2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                _cond2.Style.Fill.BackgroundColor.Index = 3;

                                //endcol = 11;
                                //ExcelAddress _formatRangeAddressB = new ExcelAddress(row + i, endcol, row + i, endcol);
                                //var _cond1B = sheet.ConditionalFormatting.AddExpression(_formatRangeAddressB);
                                //_cond1B.Formula = $"COUNTBLANK($G{row + i})=0";
                                //_cond1B.StopIfTrue = false;
                                //_cond1B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                //var color = System.Drawing.ColorTranslator.FromHtml("#E2EBD5");
                                ////_cond1B.Style.Fill.BackgroundColor.Index = -4142;
                                //_cond1B.Style.Fill.BackgroundColor.Color = color;

                                //ExcelAddress _formatRangeAddress2B = new ExcelAddress(row + i, endcol, row + i, endcol);
                                //var _cond2B = sheet.ConditionalFormatting.AddContainsBlanks(_formatRangeAddress2B);
                                //_cond2B.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                //_cond2B.Style.Fill.BackgroundColor.Index = 35;

                            }

                        }
                    }
                    else
                    {
                        // First competitor can use predefined "ekipage" range that exists on sheet
                        set.Add(klass);
                        string adress = fromRange.Address;
                        toRange = fromRange;

                    }

                    // Names, horses etc
                    var startrow = toRange.Start.Row;
                    sheet.Cells[startrow + 1, 4].Value = d.Name;
                    sheet.Cells[startrow + 2, 4].Value = d.Linforare;
                    sheet.Cells[startrow + 1, 6].Value = d.Klubb;
                    sheet.Cells[startrow + 2, 6].Value = d.Hast;

                    // The Id of the ekipage
                    sheet.Cells[startrow, 2].Value = d.Id;
                    sheet.Cells[startrow + 1, 2].Value = d.Id;
                    sheet.Cells[startrow + 2, 2].Value = d.Id;
                    sheet.Cells[startrow + 3, 2].Value = d.Id;

                    // We need the klass to get the Table name
                    var tklass = classes.Single(c => c.Name == klass);

                    startrow = toRange.Start.Row;
                    int momentIndex = 0; // ID
                    foreach (Moment moment in tklass.Moments)
                    {
                        // ID generation
                        var colnum = 8;
                        momentIndex++;
                        foreach (SubMoment submoment in moment.SubMoments)
                        {
                            // ID
                            string id = d.Id + "_" + klass + "_" + moment.Name.Replace(' ', '_') + "_" + submoment.Table.Name;
                            id = d.Id + "_" + momentIndex + "_" + submoment.Table.Name; //ID

                            var rng = sheet.Cells[startrow, colnum];
                            sheet.Names.Add(id, rng);
                            colnum++;
                        }
                        startrow++;
                    }
                    UpdateProgressBarHandler(deltagarCounter);
                    UpdateProgressBarLabel("Added " + d.Name);
                    vaulterInClassCounter[klass]++;

                    var modCounter = vaulterInClassCounter[klass] % 9;

                    // Pagebreak every 9 vaulter / class
                    if (modCounter == 0)
                    {
                        int lastRow = toRange.End.Row;
                        sheet.Row(lastRow).PageBreak = true;
                    }
                }
                results.Save();
            }

            UpdateProgressBarLabel("All vaulters added to result file");
        }
    }
}
