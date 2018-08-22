using System.Linq;
using Framework.PageObjects;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc TableHeaderRow
    /// </summary>
    [Locator(How.Sizzle, "tr:has('th')")]
    public class TableHeaderRow : TableRow
    {
        /// <summary>
        /// Get the column index for this caption (column)
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public int ColumnIndex(string caption)
        {
            int indexOf = Cells.ToList().FindIndex(c => c.Text == caption);
            return indexOf;
        }
    }
}
