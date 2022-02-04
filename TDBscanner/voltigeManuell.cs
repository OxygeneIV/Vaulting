using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using System.IO;
using Framework.PageObjects;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tests.Base;
using Viedoc.viedoc.pages.components.elements;

namespace Tests.Voltige
{





    [TestClass]
    public class VoltigeTestsManuell : ViedocTestbase
    {

        public List<CompClass> classes;


        int GetId(Dictionary<int, string> dic, String value)
        {

            foreach (var item in dic)
            {
                int i = item.Key;
                String val = item.Value;
                if (val.Equals(value)) return i;
            }

            int newIndex = dic.Count();
            dic.Add(newIndex, value);
            return newIndex;
            
        }

        [TestMethod]
        [TestCategory("VOLTIGE")]
        [Browser(false)]
        public void ReadClasses()
        {
           

            // All ints are DB Ids, not what is diaplyed in the table 1, 2, 3.1, 3.2, 4  etc

            Dictionary<int, string> _classes = new Dictionary<int, string>();
            Dictionary<int, string> _clubs = new Dictionary<int, string>();
            Dictionary<int, string> _linf = new Dictionary<int, string>();
            Dictionary<int, string> _horse = new Dictionary<int, string>();

            Dictionary<int, string> _comp = new Dictionary<int, string>();
            Dictionary<int, string> _ekipage = new Dictionary<int, string>();
            Dictionary<int, string> _classid2klassnummer = new Dictionary<int, string>();


            var lines = File.ReadAllLines("C:/privat/voltigekalkyler/voltigekalkyler1.txt");
            Array.Sort(lines);

            Trace.WriteLine("Antal rader = "+lines.Count());

            

            foreach (string line in lines)
            {
                
                List<String> row = line.Split(';').ToList();
                // Klass 5 Svår klass lag junior, 3 - 9 st; Stockholms Voltigeförening; -; -; Team Safir; Elda Lindberg; Tilda Stubbans; Signe Zetterberg; Agnes Stening; Minna Ehinger; Maja Jordansson Pinto; Nelly Lindhé;
                String klassnummer = row[0].Trim();
                String classs = row[1].Trim();
                String club = row[2].Trim();
                String linf = row[3].Trim();
                String horse = row[4].Trim();
                String note = row[5].Trim();
                List<String> comp = row.Skip(6).ToList();
                comp.RemoveAll(p => p.Trim().Length < 1);

                int classId = GetId(_classes, classs);
                int clubId = GetId(_clubs, club);
                int _linfId = GetId(_linf, linf);
                int horseId = GetId(_horse, horse);

                String ekip = String.Join("|", classId.ToString(), klassnummer, classs,
                               _linfId.ToString(), linf,
                               horseId.ToString(), horse,
                               clubId.ToString(), club,
                               note);


                String competitorstring = "";
                foreach(String competitor in comp)
                {
                   int compId = GetId(_comp, competitor.Trim());
                    ekip = ekip + "|" + compId.ToString() + "|" + competitor;
                }

                int ekipId = GetId(_ekipage, ekip);

                if(!_classid2klassnummer.ContainsKey(classId))
                    _classid2klassnummer.Add(classId, klassnummer);

            }


            foreach (KeyValuePair<int, string> kvp in _classes)
            {
                Trace.WriteLine(kvp.Key + "|" + _classid2klassnummer[kvp.Key] +"|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _linf)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _horse)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _clubs)
            {
                string country = "SE";
                string club = kvp.Value;
                string clubLowerCase = club.ToLower();

                switch (clubLowerCase)
                {
                    // Check country
                    case "denmark":
                        club = "??";
                        country = "DK";
                        break;
                    case "norway":
                        club = "??";
                        country = "NO";
                        break;
                    case "finland":
                        club = "??";
                        country = "FI";
                        break;
                    case "okänd klubb":
                        club = "??";
                        country = "??";
                        break;
                }

                Trace.WriteLine(kvp.Key + "|" + club + "|" + country);
            }
            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _comp)
            {
                Trace.WriteLine(kvp.Key + "|" + kvp.Value);
            }


            Trace.WriteLine("---------------------------");
            foreach (KeyValuePair<int, string> kvp in _ekipage)
            {
                Trace.WriteLine(kvp.Value);
            }
            Trace.WriteLine("---------------------------");


        }
    }
}


