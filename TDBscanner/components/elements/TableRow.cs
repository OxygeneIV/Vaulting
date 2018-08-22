using System;
using System.Collections.Generic;
using System.Linq;
using Framework.PageObjects;

namespace Viedoc.viedoc.pages.components.elements
{
    [Locator(How.Sizzle, "tr:has('td')")]
    public class TableRow : PageObject
    {
        public IEnumerable<TableCell> Cells { get; set; }

        /// <summary>
        /// The delegate set by Table, helps us find the column index
        /// </summary>
        public Func<string, int> GetColumnIndex { get; set; }

  
        /// <summary>
        ///  Get the cell for this column
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public TableCell GetCell(string caption)
        {
            int index = GetColumnIndex_(caption);
            return Cells.ToList()[index];
        }

        /// <summary>
        /// Get the cell text for this column
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public string GetCellText(string caption)
        {
            return GetCell(caption)
                .Text;
        }

        private int GetColumnIndex_(string caption)
        {
            return GetColumnIndex(caption);
        }
    }
}
