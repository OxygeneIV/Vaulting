using System;
using System.Collections.Generic;
using System.Linq;
using Framework.PageObjects;
using Framework.WaitHelpers;
using OpenQA.Selenium;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc Table element
    /// </summary>
    public class Table : Table<TableHeaderRow<TableCell>, TableRow<TableCell>, TableCell>
    {

    }


    /// <summary>
    /// Generic Table
    /// </summary>
    [Locator(How.Tag, "table")]
    public class Table<THeaderRow, TRow, TCell> : PageObject
        where THeaderRow : TableHeaderRow<TCell>, new()
        where TRow : TableRow<TCell>, new()
        where TCell : TableCell, new()
    {
        // Keep track of headers in the table
        protected readonly Dictionary<Enum, string> headerDictionary_ = new Dictionary<Enum, string>();

        public int GetColumnIndex(string caption)
        {
            return GetColumnIndex_(caption);
        }


        public List<TRow> GetRows(string caption, string cellValue,bool contains = false)
        {
            Log.Info("Table - GetRows Start");
            Wait.UntilOrThrow(() => Displayed, message: "Check table is visible...");

            int columnIndex = GetColumnIndex(caption);
            int tdIndex = columnIndex + 1;

            string locator = null;

            locator = !contains ? $".//tr[td[{tdIndex}][.//text()='{cellValue}']]" : 
                                  $".//tr[td[{tdIndex}][contains(text(),'{cellValue}')]]";

            var elements = GetWrappedElement().FindElements(By.XPath(locator));
            int i = 0;

            var pageObjectElements = elements.Select(e =>
                {
                    var po = PageObjectFactory.CreatePageObject<TRow>(this, e);
                    po.GetColumnIndex = GetColumnIndex_;
                    po.ColumnHeader = ColumnHeader;
                    po.GetHeaderRow = GetHeaderRow;
                    po.Index = i;
                    po.GetRowWithIndex = GetRowWithIndex;
                    i++;
                    return po;
                }
            ).ToList();

            return pageObjectElements;
        }

        public TRow GetRowWithIndex(int i)
        {
            return Rows.Single(r=>r.Index == i);
        }

        [Locator("caption")]
        protected Label caption_ = null;

        public string Caption => caption_.Text;

        /// <summary>
        /// Get the rows and add delegate to help us find correct column
        /// </summary>
        public virtual IEnumerable<TRow> Rows
        {
            get
            {
                Log.Info("Table - GetRows Start");
                Wait.UntilOrThrow(() =>Displayed,message:"Check table is visible...");
                List<TRow> newList = new List<TRow>();
                int i = 0;
                foreach (var row in Rows_)
                {
                    row.GetColumnIndex = GetColumnIndex_;
                    row.ColumnHeader = ColumnHeader;
                    row.GetHeaderRow = GetHeaderRow;
                    row.Index = i;
                    row.GetRowWithIndex = GetRowWithIndex;
                    
                    newList.Add(row);
                    i++;
                }
                var rr = newList.Cast<TRow>();
                Log.Info("Table - GetRows End");
                return rr;
            }
        }

        public IList<string> ColumnNames
        {
            get
            {
                return HeaderRow_.Cells.Select(s => s.Text).ToList();
            }
        }

        protected IEnumerable<TRow> Rows_ { get; set; }

        protected THeaderRow HeaderRow_ { get; set; }

        /// <summary>
        /// Get row having text cellText in column caption
        /// </summary>
        /// <param name="caption"></param>
        /// <param name="cellText"></param>
        /// <returns></returns>
        public TRow GetRow(string caption, string cellText)
        {
            Wait.UntilOrThrow(() => Displayed, 10, 500, "Wait for table to be displayed");
            int columnIndex = GetColumnIndex_(caption);
            return GetRow(columnIndex, cellText);
        }

        public TRow GetRow(int columnIndex, string cellText, bool contains = false)
        {
            Wait.UntilOrThrow(() => Displayed, 10, 500, "Wait for table to be displayed");
            if (contains)
            {
                return Rows.First(row => row.Cells.ToList()[columnIndex].Text.Contains(cellText));
            }

            return Rows.Single(row => row.Cells.ToList()[columnIndex].Text == cellText);
        }


        public TRow GetRow(Enum header, string value)
        {
            var row = GetRow(ColumnHeader(header), value);
            return row;
        }

        public int GetRowIndex(string caption, string cellText)
        {
            int columnIndex = GetColumnIndex_(caption);
            return Rows.ToList().FindIndex(row => row.Cells.ToList()[columnIndex].Text == cellText);
        }
        public int GetRowIndex(Enum header, string value)
        {
            var row = GetRowIndex(ColumnHeader(header), value);
            return row;
        }

        public Table<THeaderRow, TRow, TCell> GetTable()
        {
            return this;
        }

        public THeaderRow GetHeaderRow()
        {
            // Make sure the headerRow is Displayed
            Wait.UntilOrThrow(() => HeaderRow_.Displayed,5,message: "(GetHeaderRow) => Wait for header row to be displayed");
            return HeaderRow_;
        }

        public bool RowExists(Enum header, string value)
        {
            return GetRowIndex(header, value) != -1;
        }

        protected int GetColumnIndex_(string caption)
        {
            Wait.UntilOrThrow(() => HeaderRow_.Displayed, 5, message: "(GetColumnIndex_) => Wait for header row displayed");
            return HeaderRow_.ColumnIndex(caption);
        }
        protected string ColumnHeader(Enum header)
        {
            return headerDictionary_[header];
        }


    }

    [Locator(How.Sizzle, "tr:has('td')")]
    public class TableRow<TCell> : PageObject where TCell : TableCell, new()
    {
        public virtual IEnumerable<TCell> Cells { get; set; }

       // public Table table;

        public int Index;

        /// <summary>
        /// The delegate set by Table, helps us find the column index
        /// </summary>
        public Func<string, int> GetColumnIndex { get; set; }

        public Func<Enum, string> ColumnHeader { get; set; }

        public Func<TableHeaderRow<TCell>> GetHeaderRow { get; set; }
        public Func<int, object> GetRowWithIndex { get; internal set; }


        /// <summary>
        ///  Get the cell for this column
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public virtual TCell GetCell(string caption)
        {
            var index = GetColumnIndex(caption);
            return Cells.ToList()[index];
        }

        /// <summary>
        ///  Get the cell for this column
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public virtual TCell GetCell(int index)
        {
            return Cells.ToList()[index];
        }

        public TCell GetCell(Enum caption)
        {
            var headerText = ColumnHeader(caption);
            return GetCell(headerText);
        }

        /// <summary>
        /// Get the cell text for this column
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public string GetCellText(string caption)
        {
            return GetCell(caption).Text;
        }

        /// <summary>
        /// Get the cell text for this column
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public string GetCellText(Enum caption)
        {
            return GetCell(caption).Text;
        }
    }

    [Locator(How.Sizzle, "tr:has('th')")]
    public class TableHeaderRow<TCell> : TableRow<TCell> where TCell : TableCell, new()
    {
        /// <summary>
        /// Get the column index for this caption (column)
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public int ColumnIndex(string caption)
        {
            Log.Info($"Getting column index of column header {caption}");
            var theCells = Cells.ToList();
            int indexOf = -1;

            Log.Info($"Getting the header cells => {theCells.Count}");
             Wait.UntilOrThrow(() =>
            {
                //indexOf = Cells.ToList().FindIndex(c => c.Text == caption);
                indexOf = theCells.FindIndex(c => c.Text == caption);
                return true;
            },20,1000,"Waiting to find the header cell texts");

            //var indexOf = theCells.FindIndex(c => c.Text == caption);
            Log.Info($"Getting column index of column header returned {indexOf}");
            if (indexOf == -1)
            {
                Log.Info("Got these texts..");
                foreach (var c in theCells) Log.Debug(c.Text);
            }
            return indexOf;
        }


    }
}
