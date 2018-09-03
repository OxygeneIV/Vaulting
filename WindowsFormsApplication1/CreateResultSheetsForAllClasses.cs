using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace WindowsFormsApplication1
{
    partial class Form1
    {

        private void backgroundWorkerCreateClassResultsSheets_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            this.CreateClassResultSheets();
        }

        private void backgroundWorkerCreateClassResultsSheets_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
                showMessageBox("Created result sheets for all classes");
            }
        }

        private void buttonCreateResultSheets_Click(object sender, EventArgs e)
        {
            backgroundWorkerCreateClassResultsSheets.RunWorkerAsync();
            bool hasAllThreadsFinished = false;
            while (!hasAllThreadsFinished)
            {
                hasAllThreadsFinished = backgroundWorkerCreateClassResultsSheets.IsBusy == false;
                Application.DoEvents(); //This call is very important if you want to have a progress bar and want to update it
                                        //from the Progress event of the background worker.
                System.Threading.Thread.Sleep(50);     //This call waits if the loop continues making sure that the CPU time gets freed before
                                                       //re-checking.
            }
        }


        private Image DrawTextImage(String currencyCode, Font font, Color textColor, Color backColor)
        {
            return DrawTextImage(currencyCode, font, textColor, backColor, Size.Empty);
        }

        private Image DrawTextImage(String currencyCode, Font font, Color textColor, Color backColor, Size minSize)
        {
            //first, create a dummy bitmap just to get a graphics object
            SizeF textSize;
            using (Image img = new Bitmap(1, 1))
            {
                using (Graphics drawing = Graphics.FromImage(img))
                {
                    //measure the string to see how big the image needs to be
                    textSize = drawing.MeasureString(currencyCode, font);
                    if (!minSize.IsEmpty)
                    {
                        textSize.Width = textSize.Width > minSize.Width ? textSize.Width : minSize.Width;
                        textSize.Height = textSize.Height > minSize.Height ? textSize.Height : minSize.Height;
                    }
                }
            }

            //create a new image of the right size
            Image retImg = new Bitmap((int)textSize.Width, (int)textSize.Height);
            using (var drawing = Graphics.FromImage(retImg))
            {
                //paint the background
                drawing.Clear(backColor);

                //create a brush for the text
                using (Brush textBrush = new SolidBrush(textColor))
                {
                    drawing.DrawString(currencyCode, font, textBrush, 0, 0);
                    drawing.Save();
                }
            }
            return retImg;
        }

        /// <summary>
        /// create the result files for classes in startlist
        /// </summary>
        private void CreateClassResultSheets()
        {
            var classes = readClasses();
            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(classes.Count);
            UpdateProgressBarLabel("");

            FileInfo resultat = new FileInfo(resultfile);
            //Image im = new Bitmap(ridsportlogo);
            //Image imVoid = new Bitmap(logovoid);

            //Image prel = new Bitmap(preliminaryResults);
            string reference = "ResultTemplate";
            int classCount = 0;
            int maxcount = classes.Count;


            // First delete all sheets
            using (var results = new ExcelPackage(resultat))
            {
                foreach (Klass klass in classes)
                {
                    string className = klass.Name;

                        // Delete old sheets
                        if (results.Workbook.Worksheets.Any(s => s.Name == className))
                        {
                            results.Workbook.Worksheets.Delete(className);
                        }
                }

                results.Save();
            }

            using (var results = new ExcelPackage(resultat))
            {

                // Copy main sheet to each class results and update header
                foreach (Klass klass in classes)
                {


                    string className = klass.Name;

                    // SM & NM
                    if (className.EndsWith(".2"))
                        continue;


                        try
                        {
                        classCount++;

                        var ws = results.Workbook.Worksheets;
                        ws.First().Select();

                        //UV, use special templates
                        reference = klass.ResultTemplate;

                        var classWorksheet = ws.Copy(reference, className);

                        var refRange = classWorksheet.Cells["ekipage"];
                        var start = refRange.Start;
                        var end = refRange.End;
                        var cell = classWorksheet.Cells[start.Row,12];


                            
                            ExcelHeaderFooterText t22 = classWorksheet.HeaderFooter.OddHeader;
                            t22.CenteredText = "&\"Arial,bold\"&16" + "Klass " + klass.Name + "  -  " + klass.Description + "&B" + "&\"Arial\"&8" + (char)13 + "&P (&N)";
                            //t22.LeftAlignedText = "SM/NM 2018, Laholm/Caprifolen   ";
                            //t22.InsertPicture(im, PictureAlignment.Left);
                            //t22.InsertPicture(imVoid, PictureAlignment.Right);
                            

                        // Type the moments that are defined for the class
                        int counter = 0;

                        string totalJudge = "";

                        

                        foreach (Moment mom in klass.Moments)
                        {
                            counter++;
                            classWorksheet.Cells[$"round{counter}"].Value = mom.Name;
                            string judgetext = string.Format("{0,15}", mom.Name);
                            // Can be calculated, but not yet...
                            int subcounter = 8;
                            
                            foreach (SubMoment submom in mom.SubMoments)
                            {
                                string judgeName = string.Format("{0,-40}", submom.Table.judge.Fullname);

                                judgetext = judgetext + "   " + submom.Table.Name + ": " + judgeName;

                            }
                            totalJudge = totalJudge + judgetext + "\n";

                          

                            foreach (SubMoment submom in mom.SubMoments)
                            {
                                int row = classWorksheet.Cells[$"round{counter}"].Start.Row;
                                classWorksheet.Cells[row, subcounter].Value = submom.Name;
                                subcounter++;                            
                            }
                        }

                        // just some test
                        //
                        totalJudge = totalJudge + "\n\n";
                        var img = DrawTextImage(totalJudge, new Font("Lucida Console", 9, FontStyle.Italic), Color.Black, Color.White);

                        ExcelHeaderFooterText t5 = classWorksheet.HeaderFooter.OddFooter;
                        t5.InsertPicture(img, PictureAlignment.Centered);


                        classWorksheet.PrinterSettings.RepeatRows = new ExcelAddress(className+"!1:6");
                        UpdateProgressBarLabel("Added result sheet for class " + className);
         
                    }
                    catch (Exception e)
                    {
                        UpdateMessageTextBox($"Failed to clone template sheet {reference} to class {className}: " + e.Message);                        
                        UpdateProgressBarLabel($"Failed to clone template sheet {reference}: " + e.Message);
                    }

                    UpdateProgressBarHandler(classCount);


        }

                results.Save();
            }

            UpdateProgressBarLabel("All result sheets created");
        }
    }
}
