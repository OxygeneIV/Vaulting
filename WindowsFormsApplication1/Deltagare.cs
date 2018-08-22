using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class Deltagare
    {
        public string Klass;
        public string Name;
        public string Linforare;
        public string Klubb;
        public string Hast;
        public string Id;

        public Deltagare Duplicate()
        {
            Deltagare d = new Deltagare();
            d.Name = Name;
            d.Klass = Klass;
            d.Linforare = Linforare;
            d.Klubb = Klubb;
            d.Hast = Hast;
            d.Id = Id;
            return d;
        }


        public static Deltagare RowToClass(ExcelRange range)
        {
            Deltagare delt = new Deltagare();

            delt.Klass = range.ElementAt(0).Text;
            delt.Name = range.ElementAt(1).Text;
            delt.Linforare = range.ElementAt(2).Text;
            delt.Klubb = range.ElementAt(3).Text;
            delt.Hast = range.ElementAt(4).Text;
            delt.Id = range.ElementAt(5).Text;
            return delt;
        }
    }
}
