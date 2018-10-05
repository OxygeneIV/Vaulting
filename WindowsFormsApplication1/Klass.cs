using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{

    public class Table
    {
        public string Name;
        public Judge judge;
    }

    public class Judge
    {
        public string Fullname { get; set;}
        //string fullname { get { return firstName + " " + lastname;} }
    }


    public class SubMoment
    {
        public string Name;
        public Table Table; 
    } 

    // grund kür etc
    public class Moment
    {
        public string Name;
        public string reference;
        public int Type;
        public List<SubMoment> SubMoments = new List<SubMoment>();
    }

    public class Klass
    {
        public string Name;
        public string Description;
        public string ResultTemplate;
        public int Id;
        public List<Moment> Moments = new List<Moment>();


        public static Klass RowToClass(ExcelRange range)
        {

            
            Klass klass = new Klass();
          
            
            klass.Name = range.ElementAt(0).Text.Trim();
            klass.Description = range.ElementAt(1).Text.Trim();
            klass.ResultTemplate = range.ElementAt(2).Text.Trim();
            var length = range.Columns;
            for (int i = 3; i < 7; i++)
            {
                var momentText = range.ElementAt(i).Text.Trim();
                if (string.IsNullOrEmpty(momentText) || string.IsNullOrWhiteSpace(momentText))
                    continue;

                Moment moment = new Moment();

                moment.Name = momentText;

                var submomentsText = range.ElementAt(i+4).Text;

                //judges
                var submomentsJudgesText = range.ElementAt(i + 8).Text;
                var submomentsJudges = submomentsJudgesText.Split(',');

                var submoments = submomentsText.Split(',');
                int count = 65;
                int judgeindex = 0;

                foreach (var s in submoments.ToList().GetRange(0, submomentsJudges.Count()))
                {
                    if (string.IsNullOrEmpty(s) || string.IsNullOrWhiteSpace(s))
                        continue;

                    SubMoment submoment = new SubMoment();
                    submoment.Name = s;

                    Table table = new Table();
                    table.Name = ((char)count).ToString().ToUpper();
                    count++;
                    Judge judge = new Judge();
                    judge.Fullname = submomentsJudges[judgeindex];
                    table.judge = judge;
                    submoment.Table = table;
                    moment.SubMoments.Add(submoment);
                    judgeindex++;
                }
                klass.Moments.Add(moment);
            }
            return klass;

        }

        internal static Klass RowToClass(List<string> r)
        {
            Klass klass = new Klass();


            klass.Name = r[0];
            klass.Description = r[1];
            klass.ResultTemplate = r[2];

 

               
  
            for (int i = 3; i < 7; i++)
            {
                var momentText = r[i];
                if (string.IsNullOrEmpty(momentText) || string.IsNullOrWhiteSpace(momentText))
                    continue;


               // Remove any numbers from moment text
                momentText = momentText.Replace("1", "").Replace("2", "").Trim();

                Moment moment = new Moment();
                moment.Name = momentText;
                var submomentsText = r[i + 4];

                //judges
                var submomentsJudgesText = r[i + 8];
                var submomentsJudges = submomentsJudgesText.Split(',');

                var submoments = submomentsText.Split(',');
                int count = 65;
                int judgeindex = 0;

               foreach (var s in submoments.ToList())
                {
                    if (string.IsNullOrEmpty(s) || string.IsNullOrWhiteSpace(s))
                        continue;

                    SubMoment submoment = new SubMoment();
                    submoment.Name = s;

                    Table table = new Table();
                    table.Name = ((char)count).ToString().ToUpper();
                    count++;
                    
                    
                    Judge judge = new Judge();
                    if (judgeindex < submomentsJudges.Count())
                    {
                        judge.Fullname = submomentsJudges[judgeindex];
                    }else
                    {
                        judge.Fullname = "";
                    }
                    table.judge = judge;
                    submoment.Table = table;
                    
                    moment.SubMoments.Add(submoment);
                    judgeindex++;
                }
                klass.Moments.Add(moment);
            }
            return klass;

        }
    }
}
