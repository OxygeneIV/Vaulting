using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace HorsePoints
{
    public partial class Form1 : Form
    {
        //ws.Name=="1" || ws.Name == "2" || ws.Name == "9" || ws.Name == "10" || ws.Name == "11" || ws.Name == "12"
        public static List<string> teamClasses = new List<string> {"1","2","9","10","11","12","13"};

        public Form1()
        {
            InitializeComponent();
        }

        

        public class Horse : IComparable<Horse>
        {
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

        private void btnCalculateHorsePoints_Click(object sender, EventArgs e)
        {


            var resultfile = Path.Combine(label1.Text,"sortedresults.xlsx");
            var horsefile = Path.Combine(label1.Text,@"horses.xlsx");

            if(!File.Exists(resultfile))
            {
                label2.Text = "Status : No sortedresults.xlsx found in " + label1.Text;
                return;
            }

            label2.Text = "Status : Sortedresults.xlsx was found in " + label1.Text;

            FileInfo resultat = new FileInfo(resultfile);
            FileInfo horsefileInfo = new FileInfo(horsefile);

            List<string> horses = new List<string>();

            List<Horse> definedHorses = new List<Horse>();
            List<Horse> definedHorsesTeam = new List<Horse>();
            List<Horse> definedHorsesInd = new List<Horse>();

            using (ExcelPackage results = new ExcelPackage(resultat))
            {
                try
                {
                    
                    int startsheet = 5;
                    int sheetcount = results.Workbook.Worksheets.Count();
                    for (int i = startsheet; i <= sheetcount; i++)
                    {
                        int startRow = 7;
                        ExcelWorksheet ws = results.Workbook.Worksheets[i];
                        var maxrow = ws.Dimension.End.Row;

                        int ekipages = (maxrow - startRow + 1) / 4;

                        for (int ekipage = 0; ekipage < ekipages; ekipage++)
                        {
                            var currentStartRow = startRow + (ekipage*4);
                            var horsename = ws.Cells[currentStartRow+2,6].Value.ToString();
                            horses.Add(horsename);

                            if(!definedHorses.Any(h=>h.Name == horsename))
                            {
                                Horse h1 = new Horse();
                                h1.Name = horsename;
                                definedHorses.Add(h1);
                            }

                            // TEAM
                            if (teamClasses.Contains(ws.Name)) 
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
                                if(momenttext.Length > 1) // we may have points
                                {
                                    var point = ws.Cells[currentStartRow + arow, 8].GetValue<float>();
                                    if(point > 0)
                                    {
                                        curhorse.Points.Add(point);

                                        // TEAM
                                        if (teamClasses.Contains(ws.Name))// (ws.Name == "1" || ws.Name == "2" || ws.Name == "9" || ws.Name == "10" || ws.Name == "11" || ws.Name == "12")
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
                    label2.Text = label2.Text + (char)(10) + ex.Message;
                }
                finally
                {

                }
            }

            var all = horses.Distinct().ToList();
            all.RemoveAll(s => s == "A4");
            definedHorses.RemoveAll(h=>h.Name== "A4");
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
                    
                    foreach(Horse h in definedHorses)
                    {
                        row = row + 1;
                        sheet.Cells[row, 1].Value = h.Name;
                        sheet.Cells[row, 3].Value = h.Max;
                        sheet.Cells[row, 2].Value = h.Average;
                        for (int i = 0; i < h.Points.Count; i++)
                        {
                            sheet.Cells[row, 4+i].Value = h.Points[i];
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

                }
                catch (Exception ex)
                {
                    label2.Text = label2.Text + (char)(10) + ex.Message;
                }
                finally
                {
                    results.Save();
                }
            }

            label2.Text = label2.Text + (char)(10) + "Horses.xlsx created" ;

            label2.Refresh();
        }

        private void chooseFolder_Click(object sender, EventArgs e)
        {
            ChooseFolder();
        }

        public void ChooseFolder()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
