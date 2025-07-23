using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Configuration;

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


		public class ResultObject
		{
			public string clazz;
			public string description;
			public int rank;
			public string name;
			public string horse;
      public int klassRank;

			public string toFileStyle()
			{
				var l = new List<String>() {name, clazz, description, rank.ToString(), horse};
				return String.Join("|", l);
			}
		}

		public class omvandStartordningsClass
		{
			public Klass klass;
			public Int32 max;
      public Int32 omvandRank;

		}


  //      public void extractFromSortedFile_org()
		//{

		//	List<string> omvandsklass = new List<string>();
		//	List<int> maxPerClass = new List<int>();
		//	var classes = readClasses();

		//    //var escamilo = ConfigurationManager.AppSettings["escamilo"];

  //          var omvandclasses = ConfigurationManager.AppSettings["omvandclasses"].Split(',').Select(s => s.Trim()).ToList();
		//	var maxomvandclasses = ConfigurationManager.AppSettings["maxomvandclasses"].Split(',').Select(s => s.Trim()).ToList();

		//	omvandsklass.AddRange(omvandclasses);
		//	maxPerClass.AddRange(maxomvandclasses.Select(s=>Int32.Parse(s)));

		//	List< omvandStartordningsClass> omvandStartordningsClasses = new List<omvandStartordningsClass>();

		//	for (int i = 0; i < omvandclasses.Count(); i++)
		//	{
		//		omvandStartordningsClass o = new omvandStartordningsClass();
		//		o.klass = classes.Single(c => c.Name == omvandclasses[i]);
		//		o.max = maxPerClass[i];
		//		omvandStartordningsClasses.Add(o);
		//	}


		//	String extracted = omvandfile + ".txt";
		//	if (File.Exists(sortedresultsfile))
		//	{
		//		File.Delete(omvandfile);
		//		File.Copy(sortedresultsfile, omvandfile);
		//	}
		//	else
		//	{
		//		UpdateMessageTextBox("Need sorted results for omvänd startordning...");
		//		return;
		//	}

		//	if (File.Exists(extracted))
		//	{
		//		File.Delete(extracted);
		//	}

		//	// var classes = readClasses();


		//	var max = classes.Count();

		//	UpdateProgressBarHandler(0);
		//	UpdateProgressBarMax(max);
		//	UpdateProgressBarLabel("");
		//	UpdateProgressBarLabel("Starting Result Extract!!");
		//	UpdateMessageTextBox("omvänd startordning...");
		 

		//	var MyApp = new Application();
		//	MyApp.Visible = false;
		//	var workbooks = MyApp.Workbooks;
		//	var MyBook = workbooks.Open(omvandfile);

		//	int counter = 0;
			
		  
		//	List<ResultObject> goodPeople = new List<ResultObject>();
		//    List<ResultObject> eliminated = new List<ResultObject>();

  //          foreach (omvandStartordningsClass klass in omvandStartordningsClasses)
		//	{

		//	    string className = klass.klass.Name;
		//	    var MySheet = MyBook.Sheets[className];

		//	    MySheet.Activate();
		//	    UpdateMessageTextBox($"Looking at {className}");

  //              int startrow = 7;
		//		int rank = 0;
  //              while (true) // still readable
  //              {
  //                  //if (rank < klass.max)
  //                  //{ 
  //                  counter++;
				

		//			int namerow = startrow + 1;
		//			int horserow = startrow + 2;
		//			if (MySheet.Cells[namerow, 4].Value2 != null)
		//			{

		//				string name = MySheet.Cells[namerow, 4].Value.ToString();
		//				string horse = MySheet.Cells[horserow, 6].Value.ToString();
		//			    //if (escamilo == "1" && horse.ToLower().Contains("escamilo"))
		//			    //{
		//			    //    horse = horse.Replace("  ", " ");
		//			    //}

		//			    rank = rank + 1;

		//				ResultObject r = new ResultObject();
		//				r.clazz = className;
		//				r.description = klass.klass.Description;
		//				r.horse = horse;
		//				r.name = name;
		//				r.rank = rank;

	
		//				File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);

		//			    if (rank <= klass.max)
		//			    {
		//			        goodPeople.Add(r);
		//			    }
		//			    else
  //                      {
  //                          eliminated.Add(r);
  //                      }
		//				startrow = startrow + 4;
		//			}
		//			else
		//			{
		//			    File.AppendAllText(extracted, "NO more competitors in Class " + className + " startrow = " + startrow + Environment.NewLine);
		//			    break;
		//			}
		//		//}

		//	    }
  //          }

		//	MyBook.Close(true);
		//	workbooks.Close();
		//	MyApp.Quit();

		//	Marshal.ReleaseComObject(MyBook);
		//	Marshal.ReleaseComObject(workbooks);
		//	Marshal.ReleaseComObject(MyApp);
		//	MyBook = null;
		//	workbooks = null;
		//	MyApp = null;
		//	counter = 0;
	
		//	UpdateMessageTextBox("Reading SortedResults completed , got PASSED=" + goodPeople.Count() +
		//	                     "   ELIMINATED=" + eliminated.Count());

		//    File.AppendAllText(extracted, "Reading SortedResults completed, got PASSED="+goodPeople.Count() +
		//                                  "   ELIMINATED="+ eliminated.Count() + Environment.NewLine);

  //          List<String> horsenames = new List<string>();

  //          Dictionary<String, List<ResultObject>> horseVsVoltigor = new Dictionary<String, List<ResultObject>>();

		//	while (goodPeople.Count > 0)
		//	{
		//		File.AppendAllText(extracted, "Main Loop - Got " + goodPeople.Count + "  competitors" +Environment.NewLine);

		//		// Cleanup 

		//		foreach (omvandStartordningsClass startordningsClass in omvandStartordningsClasses)
		//		{
		//			if (goodPeople.Count() == 0)
		//				break;

		//			File.AppendAllText(extracted, "Using currently top ranked from klass " + startordningsClass.klass.Name + Environment.NewLine);
		//			Boolean stillThere = goodPeople.Any(p => p.clazz == startordningsClass.klass.Name);

		//			if (!stillThere)
		//			{
		//				UpdateMessageTextBox("No more in that class, go to next");
		//				continue;
		//			}

		//			ResultObject r = goodPeople.First(p => p.clazz == startordningsClass.klass.Name);

		//			String horse = r.horse;
		//		    File.AppendAllText(extracted,
		//		        "Selected top ranked competitor from klass " + startordningsClass.klass.Name + " = " +
		//		        r.toFileStyle() + Environment.NewLine);
  //                  File.AppendAllText(extracted, "Selected horse from currently top ranked in class "+ startordningsClass.klass.Name + " = " + horse + Environment.NewLine);
		//			UpdateMessageTextBox("Horse : " + horse);
		//			horsenames.Add(horse);
		//			List<ResultObject> removable = goodPeople.FindAll(p => p.horse == horse);
		//			File.AppendAllText(extracted, "Got totally " + removable.Count + " competitors with that horse " + Environment.NewLine);

  //                  horseVsVoltigor[horse] = removable;


  //                  int g = removable.Count;
  //                  // The rest
  //                  for (int i = 0; i < g; i++)
		//			{
  //                      ResultObject people = removable[i];
  //                      File.AppendAllText(extracted, "Removing  " + people.toFileStyle() + Environment.NewLine);
		//				goodPeople.Remove(people);
		//			}

		//			File.AppendAllText(extracted, "After removal, got " + goodPeople.Count() + " competitors left " + Environment.NewLine);
		//			UpdateMessageTextBox("After removal, got " + goodPeople.Count() + " voltigörer");
		//		}
				
			   

		//	}


		//	horsenames.Reverse();
		//	File.AppendAllText(extracted, "Final Horse reverse order..." + Environment.NewLine);
		//    UpdateMessageTextBox("Final Horse reverse order...");
  //          foreach (String hname in horsenames)
		//	{
		//	    UpdateMessageTextBox("Horse : " + hname);
  //              File.AppendAllText(extracted, hname + Environment.NewLine);

  //              // Voltigörer
  //              List<ResultObject> vresults = horseVsVoltigor[hname];
  //              var result = from m in vresults
  //                           orderby m.rank, m.clazz
  //                           select m;
  //              var resRevered = result.Reverse();

  //              File.AppendAllText(extracted,"Voltigörer med denna häst..." + Environment.NewLine);
  //              foreach (ResultObject r in resRevered)
  //              {
  //                  File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);

  //              }

  //          }

		//    File.AppendAllText(extracted, "Eliminated..." + Environment.NewLine+ Environment.NewLine);

		//    foreach (omvandStartordningsClass startordningsClass in omvandStartordningsClasses)
		//    {
		//        List<ResultObject> removable = eliminated.FindAll(p => p.clazz == startordningsClass.klass.Name);

		//        if (!removable.Any())
		//        {
		//            UpdateMessageTextBox("No one eliminated in class : " + startordningsClass.klass.Name);
		//            File.AppendAllText(extracted, "No one eliminated in class : " + startordningsClass.klass.Name + Environment.NewLine);
		//            File.AppendAllText(extracted, "" + Environment.NewLine);
  //                  continue;	            
		//        }

		//        UpdateMessageTextBox($"Got {removable.Count} eliminated from class  " + startordningsClass.klass.Name);
		//        File.AppendAllText(extracted, $"Got {removable.Count} eliminated from class  " + startordningsClass.klass.Name + Environment.NewLine);
  //              foreach (ResultObject r in removable)
		//        {	          
		//            UpdateMessageTextBox(r.toFileStyle());
		//            File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);
		            
		//        }
		//        File.AppendAllText(extracted, "" + Environment.NewLine);
  //          }



		//    File.AppendAllText(extracted, "END of reverse order calc..." + Environment.NewLine);
		//    UpdateMessageTextBox("END of reverse order calc...");

  //      }

        public void extractFromSortedFile()
        {

            List<string> omvandsklass = new List<string>();
            List<int> maxPerClass = new List<int>();
            var classes = readClasses();

            //var escamilo = ConfigurationManager.AppSettings["escamilo"];

            var omvandclasses = ConfigurationManager.AppSettings["omvandclasses"].Split(',').Select(s => s.Trim()).ToList();
            var maxomvandclasses = ConfigurationManager.AppSettings["maxomvandclasses"].Split(',').Select(s => s.Trim()).ToList();

            omvandsklass.AddRange(omvandclasses);
            maxPerClass.AddRange(maxomvandclasses.Select(s => Int32.Parse(s)));

            List<omvandStartordningsClass> omvandStartordningsClasses = new List<omvandStartordningsClass>();

            for (int i = 0; i < omvandclasses.Count(); i++)
            {
                omvandStartordningsClass o = new omvandStartordningsClass();
                o.klass = classes.Single(c => c.Name == omvandclasses[i]);
                o.max = maxPerClass[i];
                o.omvandRank = i+1;
                omvandStartordningsClasses.Add(o);
            }


            String extracted = omvandfile + ".txt";
            if (File.Exists(sortedresultsfile))
            {
                File.Delete(omvandfile);
                File.Copy(sortedresultsfile, omvandfile);
            }
            else
            {
                UpdateMessageTextBox("Need sorted results for omvänd startordning...");
                return;
            }

            if (File.Exists(extracted))
            {
                File.Delete(extracted);
            }

            // var classes = readClasses();


            var max = classes.Count();

            UpdateProgressBarHandler(0);
            UpdateProgressBarMax(max);
            UpdateProgressBarLabel("");
            UpdateProgressBarLabel("Starting Result Extract!!");
            UpdateMessageTextBox("omvänd startordning...");


            var MyApp = new Application();
            MyApp.Visible = false;
            var workbooks = MyApp.Workbooks;
            var MyBook = workbooks.Open(omvandfile);

            int counter = 0;


            List<ResultObject> goodPeople = new List<ResultObject>();
            List<ResultObject> eliminated = new List<ResultObject>();

            foreach (omvandStartordningsClass klass in omvandStartordningsClasses)
            {

                string className = klass.klass.Name;
                int klassRank = klass.omvandRank;
                var MySheet = MyBook.Sheets[className];

                MySheet.Activate();
                UpdateMessageTextBox($"Looking at {className}");

                int startrow = 7;
                int rank = 0;
                while (true) // still readable
                {
                    //if (rank < klass.max)
                    //{ 
                    counter++;


                    int namerow = startrow + 1;
                    int horserow = startrow + 2;
                    if (MySheet.Cells[namerow, 4].Value2 != null)
                    {

                        string name = MySheet.Cells[namerow, 4].Value.ToString();
                        string horse = MySheet.Cells[horserow, 6].Value.ToString();
                        //horse = horse.Replace("_", " ");

                        rank = rank + 1;

                        ResultObject r = new ResultObject();
                        r.clazz = className;
                        r.description = klass.klass.Description;
                        r.horse = horse;
                        r.name = name;
                        r.rank = rank;
                        r.klassRank = klassRank;


                        File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);

                        if (rank <= klass.max)
                        {
                            goodPeople.Add(r);
                        }
                        else
                        {
                            eliminated.Add(r);
                        }
                        startrow = startrow + 4;
                    }
                    else
                    {
                        File.AppendAllText(extracted, "NO more competitors in Class " + className + " startrow = " + startrow + Environment.NewLine);
                        break;
                    }
                    //}

                }
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
            counter = 0;

            UpdateMessageTextBox("Reading SortedResults completed , got PASSED=" + goodPeople.Count() +
                                 "   ELIMINATED=" + eliminated.Count());

            File.AppendAllText(extracted, "Reading SortedResults completed, got PASSED=" + goodPeople.Count() +
                                          "   ELIMINATED=" + eliminated.Count() + Environment.NewLine);

            List<String> horsenames = new List<string>();

            Dictionary<String, List<ResultObject>> horseVsVoltigor = new Dictionary<String, List<ResultObject>>();

             
            


            while (goodPeople.Count > 0)
            {

        // Hitta top
        ResultObject r = null;

        // En metod
        var result = from m in goodPeople
                             orderby m.rank, m.klassRank
                             select m;


        UpdateMessageTextBox("Omvänd select size = " + result.Count());
        r = result.First();


        // En annan metod

        //ResultObject r = null;
        //        foreach (omvandStartordningsClass klass in omvandStartordningsClasses)
        //        {
        //              if(goodPeople.Count > 0)
        //              {
        //                 var result = from m in goodPeople
        //                 where m.clazz == klass.klass.Name
        //                 orderby m.rank
        //                 select m;

        //                 if (!result.Any()) continue;

        //                  r = result.First();




                //ResultObject r = result.First();
                String horse = r.horse;
                horsenames.Add(horse);
                List<ResultObject> removable = goodPeople.FindAll(p => p.horse == horse);
                File.AppendAllText(extracted, "Got totally " + removable.Count + " competitors with that horse " + Environment.NewLine);

                horseVsVoltigor[horse] = removable;


                int g = removable.Count;
                // The rest
                for (int i = 0; i < g; i++)
                {
                    ResultObject people = removable[i];
                    File.AppendAllText(extracted, "Removing  " + people.toFileStyle() + Environment.NewLine);
                    goodPeople.Remove(people);
                }

                File.AppendAllText(extracted, "After removal, got " + goodPeople.Count() + " competitors left " + Environment.NewLine);
                UpdateMessageTextBox("After removal, got " + goodPeople.Count() + " voltigörer");


                File.AppendAllText(extracted, "Main Loop - Got " + goodPeople.Count + "  competitors" + Environment.NewLine);

      }


      horsenames.Reverse();
            File.AppendAllText(extracted, "Final Horse reverse order..." + Environment.NewLine);
            UpdateMessageTextBox("Final Horse reverse order...");
            foreach (String hname in horsenames)
            {
                UpdateMessageTextBox("Horse : " + hname);
                File.AppendAllText(extracted, hname + Environment.NewLine);

                // Voltigörer
                List<ResultObject> vresults = horseVsVoltigor[hname];
                var result = from m in vresults
                             orderby m.rank, m.clazz
                             select m;
                var resRevered = result.Reverse();

                File.AppendAllText(extracted, "Voltigörer med denna häst..." + Environment.NewLine);
                foreach (ResultObject r in resRevered)
                {
                    File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);

                }

            }

            File.AppendAllText(extracted, "Eliminated..." + Environment.NewLine + Environment.NewLine);

            foreach (omvandStartordningsClass startordningsClass in omvandStartordningsClasses)
            {
                List<ResultObject> removable = eliminated.FindAll(p => p.clazz == startordningsClass.klass.Name);

                if (!removable.Any())
                {
                    UpdateMessageTextBox("No one eliminated in class : " + startordningsClass.klass.Name);
                    File.AppendAllText(extracted, "No one eliminated in class : " + startordningsClass.klass.Name + Environment.NewLine);
                    File.AppendAllText(extracted, "" + Environment.NewLine);
                    continue;
                }

                UpdateMessageTextBox($"Got {removable.Count} eliminated from class  " + startordningsClass.klass.Name);
                File.AppendAllText(extracted, $"Got {removable.Count} eliminated from class  " + startordningsClass.klass.Name + Environment.NewLine);
                foreach (ResultObject r in removable)
                {
                    UpdateMessageTextBox(r.toFileStyle());
                    File.AppendAllText(extracted, r.toFileStyle() + Environment.NewLine);

                }
                File.AppendAllText(extracted, "" + Environment.NewLine);
            }



            File.AppendAllText(extracted, "END of reverse order calc..." + Environment.NewLine);
            UpdateMessageTextBox("END of reverse order calc...");

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
      UpdateMessageTextBox("sortedresultsfile copied...");

      // Sätt färger på cellerna

      var MyApp = new Application();
            MyApp.Visible = false;
            var workbooks = MyApp.Workbooks;
            Workbook MyBook = workbooks.Open(sortedresultsfile, ReadOnly: false);

            //MyApp = new Application();
 		        // workbooks = MyApp.Workbooks;
			      // MyBook = workbooks.Open(sortedresultsfile);
			
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


            //MyApp.Visible = false;
            //workbooks = MyApp.Workbooks;

            //MyBook = workbooks.Open(sortedresultsfile, ReadOnly: false);

            //foreach (Klass className in classes)
            //{

            //    Worksheet sss=  MyBook.Sheets[className.Name];
            //    Range r =  sss.UsedRange;
            //    int g = r.Rows.Count;


            //    var MySheet = MyBook.Sheets[className.Name];


            //   Range range2 = MySheet.UsedRange.SpecialCells(XlCellType.xlCellTypeAllFormatConditions);
            //    int hhh= range2.Cells.Count;

            //    foreach (Range c in range2.Cells)
            //    {

            //        var color = c.DisplayFormat.Interior.Color;

            //        if (color == 65280)
            //        {
            //            c.Interior.Color = 65280;
            //        }
            //        else
            //        {
            //            if (color == 14348258  || color == 14019554)
            //            {
            //                c.Interior.Color = 14348258;
            //            }
            //        }
            //        var t2 = c.Value2;
            //        var t1 = c.Value;

            //    }

            //}

            //MyBook.Save();
            //MyBook.Close();



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
