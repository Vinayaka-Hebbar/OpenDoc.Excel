using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenDoc.Excel
{
    internal abstract class RangeContainer : Container<ExcelItem, IRangeContainer>, IRangeContainer, IBaseContainer<ExcelItem>
    {
        public RangeContainer(IRangeContainer parent, bool canAppend = true, bool isElement = false) : base(parent, canAppend, isElement)
        {
            Style = parent.Style;
        }

        public RangeContainer(IRangeContainer parent, XlsStyle style, bool canAppend = true, bool isElement = false) : base(parent, canAppend, isElement)
        {
            Style = style;
        }

        public RangeContainer(IRangeContainer parent, ExcelItem item, XlsStyle style, bool canAppend = true, bool isElement = true) : base(parent, canAppend, isElement)
        {
            Style = style;
            Element = item;
        }

        public virtual ISheetContainer Sheet => Parent.Sheet;

        public override ExcelItem Element { get; } = ExcelItem.None;

        public XlsStyle Style { get; protected set; }

        public abstract XlsFlags Flags { get; }

        public XlsWorkbook Workbook
        {
            get
            {
                return Parent.Workbook;
            }
        }

        public virtual int Column => Parent.Column;

        public virtual int Row => Parent.Row;

        public virtual ExcelRangeBase GetMergedRange(int rowSpan, int colSpan)
        {
            return Parent.GetMergedRange(rowSpan, colSpan);
        }

        public abstract ExcelRangeBase GetRange();
        public abstract ExcelRangeBase GetRange(XlsStyle style);

        public abstract ExcelRangeBase GetRange(int row, int col);
        public abstract ExcelRangeBase GetRange(int row, int col, XlsStyle style);

        public abstract IRowContainer NewRow(int rowSpan);

        public abstract void SkipColumn(int colCount = 1);

        public abstract void SkipRow(int row, int rowspan);

        public abstract IRowContainer NewRow(XlsStyle style, int rowSpan = 1);

        public override bool Add(ExcelItem element)
        {
            return false;
        }

        public override bool Add(IRangeContainer container)
        {
            return false;
        }

        public virtual void SetColumn(int column = 1)
        {
            Parent.SetColumn(column);
        }
    }

    internal class CellContainer : RangeContainer, IRowContainer
    {
        private readonly IRowContainer row;
        private readonly ExcelRangeBase _cell;

        public CellContainer(IRowContainer parent, ExcelRangeBase cell) : base(parent, true, false)
        {
            _cell = cell;
            row = parent;
            Column = cell.Start.Column;
        }

        public CellContainer(IRowContainer parent, ExcelRangeBase cell, XlsStyle style) : base(parent, style, true, false)
        {
            _cell = cell;
            row = parent;
            Column = cell.Start.Column;
        }

        public override int Row => _cell.Start.Row;

        public override int Column { get; }

        public override XlsFlags Flags { get; } = XlsFlags.CellContainer;

        public override bool Add(ExcelItem element)
        {
            return false;
        }

        public override ExcelRangeBase GetMergedRange(int rowSpan, int colSpan)
        {
            return Parent.GetMergedRange(rowSpan, colSpan);
        }


        public override ExcelRangeBase GetRange()
        {
            return _cell;
        }

        public override ExcelRangeBase GetRange(XlsStyle style)
        {
            return _cell;
        }

        public override ExcelRangeBase GetRange(int row, int col)
        {
            return Parent.GetRange(row, col);
        }

        public override ExcelRangeBase GetRange(int row, int col, XlsStyle style)
        {
            return Parent.GetRange(row, col);
        }

        public override IRowContainer NewRow(int rowSpan)
        {
            _cell.RichText.Add("\n");
            return this;
        }

        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            if (_cell.IsRichText)
                _cell.RichText.Last().Text += XmlHelpers.NewLineString;
            return new CellContainer(this, _cell, style);
        }

        public override void SkipColumn(int colCount = 1)
        {
            throw new InvalidOperationException("Can't go to next column");
        }

        public override void SkipRow(int row, int rowspan)
        {
            _cell.RichText.Add(Environment.NewLine);
        }

        public IRangeContainer ApplyStyle()
        {
            throw new NotImplementedException();
        }

        public ExcelRangeBase MergedRange(int startRow, int startCol, int endRow, int endCol)
        {
            return row.MergedRange(startRow, startCol, endRow, endCol);
        }
    }

    internal sealed class TextContainer : CellContainer
    {
        public TextContainer(IRowContainer parent, ExcelRangeBase cell, XlsStyle style) : base(parent, cell, style)
        {
        }

        public override XlsFlags Flags => XlsFlags.TextContainer;
    }

    /// <summary>
    /// Container for table or div
    /// </summary>
    internal class ElementContainer : RangeContainer
    {
        public override XlsFlags Flags => Parent.Flags;
        //start column
        private int column;
        //end column
        private readonly int row;

        public ElementContainer(IRangeContainer parent, int row, int column, XlsStyle style, bool canAppend = true) : base(parent, style, canAppend: canAppend)
        {
            this.row = row;
            this.column = column;
        }

        public ElementContainer(IRangeContainer parent, bool canAppend, bool isElement) : base(parent, canAppend, isElement)
        {
            row = parent.Row;
            column = parent.Column;
        }

        public override int Row => row;

        public override int Column => column;

        public override bool Add(ExcelItem element)
        {
            return Parent.Add(element);
        }

        public override ExcelRangeBase GetRange()
        {
            if (Style.MergedCol.IsMerged)
                return GetMergedRange(Style.MergedRow.Count, Style.MergedCol.Count);
            return Parent.GetRange();
        }

        public override ExcelRangeBase GetRange(XlsStyle style)
        {
            if (Style.MergedCol.IsMerged)
                return GetMergedRange(Style.MergedRow.Count, Style.MergedCol.Count);
            return Parent.GetRange(style);
        }

        public override ExcelRangeBase GetRange(int row, int col)
        {
            return Parent.GetRange(row, col);
        }

        public override ExcelRangeBase GetRange(int row, int col, XlsStyle style)
        {
            return Parent.GetRange(row, col, style);
        }

        public override void SkipColumn(int colCount = 1)
        {
            Parent.SkipColumn(colCount);
        }

        public override void SkipRow(int row, int rowspan)
        {
            //since row count merged skip + current
            Style = Style.SetMergedRow(row, rowspan);
            Sheet.SkipRow(row, rowspan);
        }

        public override ExcelRangeBase GetMergedRange(int rowSpan, int colSpan)
        {
            return Parent.GetMergedRange(rowSpan, colSpan);
        }

        public override void SetColumn(int column = 1)
        {
            this.column = column;
        }

        public override IRowContainer NewRow(int rowSpan)
        {
            var row = Sheet.NewRow(this, Style, rowSpan);
            if (column > 1)
                row.SetColumn(column);
            return row;
        }

        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            IRowContainer row = Sheet.NewRow(this, style, rowSpan);
            if (column > 1)
                row.SetColumn(column);
            return row;
        }
    }

    internal class TemplaeContainer : RangeContainer
    {
        public override XlsFlags Flags => Parent.Flags;

        public TemplaeContainer(IRangeContainer parent) : base(parent)
        {
        }

        public TemplaeContainer(IRangeContainer parent, bool canAppend, bool isElement) : base(parent, canAppend, isElement)
        {
        }

        public override IRowContainer NewRow(int rowSpan)
        {
            return Parent.NewRow(Style, rowSpan);
        }

        public override ExcelRangeBase GetRange()
        {
            return Parent.GetRange();
        }

        public override ExcelRangeBase GetRange(XlsStyle style)
        {
            return Parent.GetRange(style);
        }

        public override ExcelRangeBase GetRange(int row, int col)
        {
            return Parent.GetRange(row, col);
        }

        public override ExcelRangeBase GetRange(int row, int col, XlsStyle style)
        {
            return Parent.GetRange(row, col, style);
        }

        public override void SkipColumn(int colCount = 1)
        {
            Parent.SkipColumn(colCount);
        }

        public override void SkipRow(int row, int rowspan)
        {
            Parent.SkipRow(row, rowspan);
        }

        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            return Parent.NewRow(style, rowSpan);
        }
    }

    internal sealed class SheetContainer : RangeContainer, ISheetContainer
    {
        private int currentRow = 1;

        private IRowContainer emptyRow;
        public IRowContainer EmptyRow
        {
            get
            {
                return emptyRow ?? (emptyRow = new RowContainer(this, -1));
            }
        }

        public SheetContainer(IRangeContainer parent, ExcelWorksheet sheet, XlsStyle style) : base(parent, style)
        {
            WorkSheet = sheet;
            Rows = new List<IRowContainer>();
        }

        public override int Row => currentRow;

        public override ISheetContainer Sheet => this;

        public IList<IRowContainer> Rows { get; }

        public IRowContainer CurrentRow
        {
            get
            {
                return Rows.LastOrDefault() ?? EmptyRow;
            }
        }

        public ExcelWorksheet WorkSheet { get; }

        public bool CanAutoFit { get; set; } = true;

        public override XlsFlags Flags { get; } = XlsFlags.SheetContainer;

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

        public override IRowContainer NewRow(int rowspan)
        {
            return NewRow(this, Style, rowspan);
        }

        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            return NewRow(this, style, rowSpan);
        }

        public IRowContainer NewRow(IRangeContainer parent, XlsStyle style, int rowspan)
        {
            // Row might be skiped
            if (Rows.Count > currentRow)
            {
                RowContainer item = new RowContainer(parent, currentRow++, style);
                Rows.Add(item);
                return item;
            }
            //todo fix understanding of row
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
            Rows.Add(row);
            return row;
        }

        public override void SetColumn(int column = 1)
        {
            //Nothing
        }

        public override void SkipColumn(int colCount = 1)
        {
            CurrentRow.SkipColumn(colCount);
        }

        public override void SkipRow(int row, int rowspan)
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
            if (Rows.Count > currentRow)
            {
                currentRow += rowCount;
                return;
            }
            foreach (var _ in Enumerable.Range(currentRow, rowCount))
            {
                currentRow++;
                Rows.Add(new RowContainer(this, currentRow));
            }
        }

        public void MakeEmptyRow()
        {
            currentRow++;
            //Top make Row
            WorkSheet.Cells[currentRow, 1].RichText.Add(XmlHelpers.SpaceString);
            Rows.Add(EmptyRow);
        }
    }

    internal sealed class RowContainer : RangeContainer, IRowContainer
    {
        internal static readonly IRowContainer Empty = new RowContainer();
        internal const int ColumnStartIndex = 1;
        private int column = ColumnStartIndex;
        readonly int totalColumns;
        int row;
        private readonly ExcelWorksheet sheet;

        private RowContainer() : base(null, ExcelItem.None, XlsStyle.None, false, false)
        {
            row = -1;
            totalColumns = -1;
        }

        public RowContainer(IRangeContainer parent, int row) : base(parent)
        {
            sheet = parent.Sheet.WorkSheet;
            this.row = row;
            totalColumns = Style.TotalColumns;
        }

        public RowContainer(IRangeContainer parent, int row, XlsStyle style) : base(parent, style)
        {
            sheet = parent.Sheet.WorkSheet;
            this.row = row;
            column = parent.Column;
            totalColumns = style.TotalColumns;
        }

        public override int Row => row;

        public override int Column => column;

        public override XlsFlags Flags { get; } = XlsFlags.RowContainer;

        public override bool Add(ExcelItem element)
        {
            //todo Cell Color for Null
            //todo analyze this
            if (element.IsNull)
                return false;
            //If First Column is Null Skip That

            if (Style.MergedCol.IsBetween(column))
            {
                column++;
                return false;
            }
            //Check Item Prev To Merged Col
            var mergedRow = Style.MergedRow;
            if (mergedRow.IsMerged)
            {
                if (mergedRow.IsBetween(Row))
                    column++;
            }
            return true;
        }

        public override ExcelRangeBase GetRange()
        {
            var range = sheet.Cells[Row, column];
            column++;
            //Make Row Border Style
            range.SetStyle(Style);
            return range;
        }

        public override ExcelRangeBase GetRange(XlsStyle style)
        {
            var range = sheet.Cells[Row, column];
            column++;
            if (range.Merge)
                return GetRange(style);
            //Make Row Border Style
            range.SetStyle(style);
            return range;
        }

        public override ExcelRangeBase GetRange(int row, int col)
        {
            column = col;
            var range = sheet.Cells[row, col];
            //Make Row Border Style
            range.SetStyle(Style);
            return range;
        }

        public override ExcelRangeBase GetRange(int row, int col, XlsStyle style)
        {
            column = col;
            var range = sheet.Cells[row, col];
            //Make Row Border Style
            range.SetStyle(style);
            return range;
        }

        public override void SetColumn(int column = 1)
        {
            this.column = column;
        }

        public override void SkipColumn(int colCount = 1)
        {
            int startCol = column;
            int endCol = column + colCount;
            column = endCol;
            var range = sheet.Cells[Row, startCol, Row, endCol];
            //Make Border Appear
            range.SetStyle(Style);
            if (range.Merge)
            {
                SkipColumn();
                return;
            }
            if (startCol != endCol)
                Style = Style.SetMergedCol(startCol, colCount);
        }

        public override void SkipRow(int row, int rowspan)
        {
            Sheet.SkipRow(row, rowspan);
        }

        public override ExcelRangeBase GetMergedRange(int rowspan, int colspan)
        {
            int startRow = Row;
            int startCol = column;
            var merge = sheet.MergedCells[startRow, startCol];
            if (merge != null)
            {
                var address = new ExcelAddress(merge);
                column += address.Columns;
                // if coulmn reaches max to next of merged row
                if (totalColumns > 0 && column >= totalColumns)
                {
                    column = Parent.Column;
                    // to next of merged row
                    var count = address.End.Row - row + 1;
                    Parent.SkipRow(row, count);
                    row += count;
                }
                return GetMergedRange(rowspan, colspan);
            }
            int rowCount = rowspan - 1;
            int colCount = colspan - 1;
            column += colspan;
            int endRow = startRow + rowCount;
            int endCol = startCol + colCount;
            ExcelRange range = sheet.Cells[startRow, startCol, endRow, endCol];
            if (startRow != endRow)
            {
                //to merge
                Style = Style.SetMergedRow(startRow, colCount);
                range.Merge = true;
            }
            else if (startCol != endCol)
            {
                Style = Style.SetMergedCol(startCol, colCount);
                range.Merge = true;
            }
            return range;

        }

        public ExcelRangeBase MergedRange(int startRow, int startCol, int rowSpan, int colSpan)
        {
            column = startCol + colSpan;
            int endRow = startRow + rowSpan - 1;
            int endCol = column - 1;
            var range = sheet.Cells[startRow, startCol, endRow, endCol];
            range.Merge = true;
            return range;
        }

        public override IRowContainer NewRow(int rowSpan = 1)
        {
            throw new NotImplementedException("Row Inside Row Not Allowed");
        }

        public IRangeContainer ApplyStyle()
        {
            var row = sheet.Row(Row);
            row.SetStyle(Style);
            return this;
        }

        public override IRowContainer NewRow(XlsStyle style, int rowSpan = 1)
        {
            throw new NotImplementedException("Row Inside Row Not Allowed");
        }
    }
}
