using OfficeOpenXml;
using System.Collections.Generic;

namespace OpenDoc.Excel
{
    public interface ISheetsContainer : IRangeContainer
    {
        IList<ISheetContainer> Sheets { get; }
        PageStyle PageStyle { get; }
    }

    public interface IRowContainer : IRangeContainer
    {
        ExcelRangeBase MergedRange(int startRow, int startCol, int rowSpan, int colSpan);
        IRangeContainer ApplyStyle();
    }

    public interface IRangeContainer : IBaseContainer<ExcelItem>, IContainerListener<ExcelItem, IRangeContainer>
    {
        ISheetContainer Sheet { get; }
        IRangeContainer Parent { get; }
        XlsFlags Flags { get; }
        XlsStyle Style { get; }
        /// <summary>
        /// start column
        /// </summary>
        int Column { get; }
        /// <summary>
        /// start Row
        /// </summary>
        int Row { get; }
        XlsWorkbook Workbook { get; }

        ExcelRangeBase GetRange(XlsStyle style);
        ExcelRangeBase GetRange();
        ExcelRangeBase GetRange(int row, int col, XlsStyle style);
        ExcelRangeBase GetRange(int row, int col);
        ExcelRangeBase GetMergedRange(int rowSpan, int colSpan);
        void SetColumn(int column = 1);
        void SkipColumn(int colCount = 1);
        void SkipRow(int row, int rowspan = 1);
        IRowContainer NewRow(XlsStyle style, int rowspan = 1);
        IRowContainer NewRow(int rowspan = 1);
    }

    public interface IExcelItem
    {

    }

    public
#if LATEST_VS
        readonly
#endif
        struct ExcelItem : IExcelItem
    {
        public static readonly ExcelItem None = new ExcelItem(true);
        public readonly bool IsNull;

        public ExcelItem(bool isNull)
        {
            IsNull = isNull;
        }
    }

    public interface ISheetContainer : IRangeContainer, IBaseContainer<ExcelItem>
    {
        IList<IRowContainer> Rows { get; }
        IRowContainer CurrentRow { get; }
        ExcelWorksheet WorkSheet { get; }
        bool CanAutoFit { get; }
        void MakeEmptyRows(int rowCount);
        void MakeEmptyRow();
        IRowContainer NewRow(IRangeContainer parent, XlsStyle style, int rowspan = 1);
        IRowContainer MakeRow(IRangeContainer parent, XlsStyle style, int row);
    }
}
