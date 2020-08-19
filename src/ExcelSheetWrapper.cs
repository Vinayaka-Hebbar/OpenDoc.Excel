using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace OpenDoc.Excel
{
    /// <summary>
    /// Sub sheet
    /// </summary>
    internal sealed class ExcelSheetWrapper : RangeContainer, ISheetContainer
    {
        private int column = 1;
        private int currentRow;
        private readonly ISheetContainer sheet;
        public ExcelSheetWrapper(IRangeContainer parent, int currentRow, XlsStyle style) : base(parent, style)
        {
            this.currentRow = currentRow;
            sheet = parent.Sheet;
            if (currentRow >= sheet.Rows.Count)
            {
                sheet.MakeEmptyRows(currentRow - sheet.Rows.Count);
            }
            Rows = new List<IRowContainer>(sheet.Rows);
        }

        private IRowContainer emptyRow;
        public IRowContainer EmptyRow
        {
            get
            {
                return emptyRow ?? (emptyRow = new RowContainer(this, -1));
            }
        }

        public IList<IRowContainer> Rows { get; }

        public IRowContainer CurrentRow
        {
            get
            {
                return Rows[currentRow - 1];
            }
        }

        public ExcelWorksheet WorkSheet => sheet.WorkSheet;

        public bool CanAutoFit => sheet.CanAutoFit;

        public override ISheetContainer Sheet => this;

        public override XlsFlags Flags => sheet.Flags;

        public override bool Add(ExcelItem element)
        {
            return true;
        }

        public override ExcelRangeBase GetMergedRange(int rowSpan, int colSpan)
        {
            return CurrentRow.GetMergedRange(rowSpan, colSpan);
        }

        public ExcelRangeBase GetRange(string name)
        {
            return WorkSheet.Cells[name];
        }

        public override ExcelRangeBase GetRange()
        {
            return CurrentRow.GetRange();
        }

        public override ExcelRangeBase GetRange(XlsStyle style)
        {
            return CurrentRow.GetRange(style);
        }

        public override ExcelRangeBase GetRange(int row, int column)
        {
            return CurrentRow.GetRange(row, column);
        }

        public override ExcelRangeBase GetRange(int row, int column, XlsStyle style)
        {
            return CurrentRow.GetRange(row, column, style);
        }

        public override IRowContainer NewRow(int rowSpan = 1)
        {
            return NewRow(this, Style, rowSpan);
        }


        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            return NewRow(this, style, rowSpan);
        }

        public IRowContainer NewRow(IRangeContainer parent, XlsStyle style, int rowspan = 1)
        {
            //Check For Prev Merged Row
            if (Rows.Count > currentRow)
            {
                RowContainer item = new RowContainer(parent, currentRow++, style);
                Rows.Add(item);
                return item;
            }
            int startRow = currentRow;
            currentRow += rowspan;
            if (rowspan == 1)
            {
                return MakeRow(parent, style, startRow);
            }
            foreach (var row in Enumerable.Range(startRow, rowspan))
            {
                Rows.Add(new RowContainer(parent, row, style));
            }
            return Rows[startRow - 1];
        }

        public IRowContainer MakeRow(IRangeContainer parent, XlsStyle style, int rowIndex)
        {
            var row = new RowContainer(parent, rowIndex, style);
            row.SetColumn(column);
            Rows.Add(row);
            return row;
        }

        public override void SetColumn(int column = 1)
        {
            //column index starts from index 0
            this.column = column;
        }

        public override void SkipColumn(int colCount = 1)
        {
            CurrentRow.SkipColumn(colCount);
        }

        public override void SkipRow(int row, int rowspan = 1)
        {
            var size = row + rowspan;
            if (Rows.Count < size)
            {
                foreach (var _ in Enumerable.Range(row, rowspan))
                {
                    currentRow++;
                    //add dummy row
                    Rows.Add(EmptyRow);
                }
            }
        }

        public void MakeEmptyRows(int rowCount)
        {
            foreach (var _ in System.Linq.Enumerable.Range(currentRow, rowCount))
            {
                currentRow++;
                Rows.Add(EmptyRow);
            }
        }

        public void MakeEmptyRow()
        {
            //Check For Prev Merged Row
            if (Rows.Count < currentRow)
            {
                Rows.Add(EmptyRow);
                return;
            }
            currentRow++;
            //Top make Row
            WorkSheet.Cells[currentRow, column].RichText.Add(XmlHelpers.SpaceString);
            Rows.Add(EmptyRow);
        }
    }
}
