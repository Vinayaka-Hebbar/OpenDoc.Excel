using OfficeOpenXml;
using OpenDoc.Utils;
using OpenDoc.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenDoc.Excel
{
    public sealed class XlsWorkbook : ITemplate<IRangeContainer>, ISheetsContainer
    {
        #region Tags
        private const string HEADERFOOTERMARGIN = "header-footer-margin";

        private const string PRINTAREA = "print-area";
        private const string PROTECTED = "protected";
        private const string COLWIDTH = "col-width";
        private const string ROWHEIGHT = "row-height";
        private const string FITTOWIDTH = "fit-to-width";
        private const string FITTOHEIGHT = "fit-to-height";
        private const string FITTOPAGE = "fit-to-page";

        private const string PAGEBREAKS = "page-breaks";
        private const string PAGEBREAKVIEW = "page-break-view";
        #endregion

        public const decimal MarginOffset = 0.39M;

        readonly XNamespace ns;

        public XlsWorkbook(IRoot root, ExcelPackage package, PageStyle style)
        {
            ExcelPackage = package;
            ExcelWorkbook = package.Workbook;
            Model = root.Model;
            ns = root.Namespace;
            Sheets = new List<ISheetContainer>();
            PageStyle = style;
            Style = new XlsStyle(style);
        }

        private int currentSheetIndex = -1;

        public ExcelPackage ExcelPackage { get; }

        public object Model { get; }

        public bool CanAppend => true;

        public bool IsElement => false;

        public IList<ISheetContainer> Sheets { get; }

        public ExcelWorksheet Worksheet => ExcelWorkbook.Worksheets[currentSheetIndex];

        public bool IsInsidesheet => false;

        public XlsWorkbook Workbook => this;

        public ISheetContainer Sheet => Sheets[currentSheetIndex];

        ExcelItem IBaseContainer<ExcelItem>.Element => ExcelItem.None;

        public IRangeContainer Parent => this;

        public ExcelWorkbook ExcelWorkbook { get; }

        public XlsStyle Style { get; }

        public PageStyle PageStyle { get; }

        public XlsFlags Flags { get; } = XlsFlags.Workbook;

        public int Column => 1;

        public int Row => 1;

        public static XlsWorkbook Create(IRoot root, XElement element)
        {
            //todo 
            var package = new ExcelPackage();

            var style = GetStyleFromAttrs(element, package.Workbook.Styles.NamedStyles.First().Style);
            return new XlsWorkbook(root, package, style);
        }

        private static PageStyle GetStyleFromAttrs(XElement element, OfficeOpenXml.Style.ExcelStyle xlsStyle)
        {
            var style = new PageStyle();
            var xlsFont = xlsStyle.Font;
            style.Font.Size = xlsFont.Size;
            style.Margin = Thickness.Zero;
            style.Font.Name = xlsFont.Name;
            style.Font.Color = System.Drawing.Color.Black;
            foreach (var attr in element.Attributes(XNamespace.Excel))
            {
                switch (attr.Name)
                {
                    case Tags.MARGIN:
                        style.Margin = Thickness.Parse(attr.Value);
                        break;
                    case Tags.FONTSIZE:
                        style.Font.Size = Convert.ToInt32(attr.Value);
                        break;
                    case Tags.FONTFAMILY:
                        style.Font.Name = attr.Value;
                        xlsFont.SetFromFont(new System.Drawing.Font(attr.Value, style.Font.Size));
                        break;
                    case Tags.HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.HorizontalAlignment = align;
                        break;
                    case Tags.VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.VerticalAlignment = valign;
                        break;
                    case Tags.COLOR:
                        style.Font.Color = PageStyle.GetColor(attr.Value);
                        break;
                }
            }
            return style;
        }

        public bool Add(ExcelItem element)
        {
            return true;
        }

        public bool Add(IRangeContainer element)
        {
            return Sheet.Add(element.Element);
        }

        public ISheetContainer Create(XElement element)
        {
            var nameAttr = element.Attribute(XNamespace.Excel, Tags.NAME);
            if (nameAttr != null)
                return Create(nameAttr.Value, element);
            return Create("Page" + Sheets.Count, element);
        }

        public ISheetContainer Create(string name, XElement element)
        {
            var style = Style;
            var margin = PageStyle.Margin;
            var sheet = ExcelWorkbook.Worksheets.Add(name);
            currentSheetIndex++;
            Thickness hfmargin = Thickness.Zero;
            PageBreak pageBreaks = PageBreak.None;
            PagePoint printArea = PagePoint.Zero;
            foreach (var attr in element.Attributes(ns))
            {
                switch (attr.Name)
                {
                    case Tags.MARGIN:
                        margin = Thickness.Parse(attr.Value);
                        break;
                    case HEADERFOOTERMARGIN:
                        hfmargin = Thickness.Parse(attr.Value);
                        break;
                    case Tags.FONTSIZE:
                        style.Font.Size = Convert.ToInt32(attr.Value);
                        break;
                    case Tags.FONTFAMILY:
                        style.Font.Name = attr.Value;
                        break;
                    case Tags.HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign halign);
                        style.SetHAlign(halign);
                        break;
                    case Tags.VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                    case PRINTAREA:
                        printArea = PagePoint.Parse(attr.Value);
                        break;
                    case PROTECTED:
                        bool.TryParse(attr.Value, out bool isProtected);
                        sheet.Protection.IsProtected = isProtected;
                        break;
                    case PAGEBREAKS:
                        pageBreaks = PageBreak.Parse(attr.Value);
                        break;
                    case PAGEBREAKVIEW:
                        bool.TryParse(attr.Value, out bool pageBreakView);
                        sheet.View.PageBreakView = pageBreakView;
                        break;
                    case COLWIDTH:
                        double colWidth;
                        colWidth = Enum.TryParse(attr.Value, out Enums.SizeType widthType) ? (sbyte)widthType : XmlConvert.ToDouble(attr);
                        if (colWidth > 0)
                            sheet.DefaultColWidth = colWidth;
                        else
                            sheet.Cells.AutoFitColumns();
                        break;
                    case ROWHEIGHT:
                        sheet.DefaultRowHeight = XmlConvert.ToDouble(attr);
                        break;
                    case FITTOPAGE:
                        sheet.PrinterSettings.FitToPage = XmlConvert.ToBoolean(attr);
                        break;
                    case FITTOWIDTH:
                        sheet.PrinterSettings.FitToWidth = XmlConvert.ToInt32(attr);
                        break;
                    case FITTOHEIGHT:
                        sheet.PrinterSettings.FitToHeight = XmlConvert.ToInt32(attr);
                        break;
                }
            }
            ExcelHelper.Apply(sheet.Cells.Style, style);
            if (margin != Thickness.Zero)
            {
                sheet.PrinterSettings.LeftMargin = (decimal)margin.Left * MarginOffset;
                sheet.PrinterSettings.TopMargin = (decimal)margin.Top * MarginOffset;
                sheet.PrinterSettings.RightMargin = (decimal)margin.Right * MarginOffset;
                sheet.PrinterSettings.BottomMargin = (decimal)margin.Bottom * MarginOffset;
            }
            if (hfmargin != Thickness.Zero)
            {
                sheet.PrinterSettings.HeaderMargin = (decimal)margin.Left * MarginOffset;
                sheet.PrinterSettings.FooterMargin = (decimal)margin.Top * MarginOffset;
            }
            if (printArea.IsGreaterThan(0))
            {
                sheet.PrinterSettings.PrintArea = sheet.Cells[printArea.Top, printArea.Left, printArea.Bottom, printArea.Right];
            }
            if (pageBreaks.Breaks != null)
            {
                foreach (var pageBreak in pageBreaks.Breaks)
                {
                    if (pageBreak.Row > 0)
                    {
                        var maxRows = sheet.Cells.Rows;
                        if (pageBreak.Col < maxRows)
                        {
                            sheet.Row(pageBreak.Row).PageBreak = true;
                        }
                    }
                    if (pageBreak.Col > 0)
                    {
                        var maxColumns = sheet.Cells.Columns;
                        if (pageBreak.Col < maxColumns)
                        {
                            sheet.Column(pageBreak.Col).PageBreak = true;
                        }
                    }
                }
                sheet.View.PageBreakView = true;
            }
            var workSheet = new SheetContainer(this, sheet, style);
            Sheets.Add(workSheet);
            return workSheet;
        }

        public ExcelRangeBase GetRange()
        {
            return Sheet.GetRange();
        }

        public ExcelRangeBase GetRange(XlsStyle style)
        {
            return Sheet.GetRange(style);
        }

        public void Dispose()
        {
            ExcelPackage.Dispose();
        }

        public ExcelRangeBase GetRange(int row, int col)
        {
            return Sheet.GetRange(row, col);
        }

        public ExcelRangeBase GetRange(int row, int col, XlsStyle style)
        {
            return Sheet.GetRange(row, col, style);
        }

        public void SkipColumn(int colCount = 1)
        {
            Sheet.SkipColumn(colCount);
        }

        public void SkipRow(int row, int rowspan)
        {
            Sheet.SkipRow(row, rowspan);
        }

        public ExcelRangeBase GetMergedRange(int rowSpan, int colSpan)
        {
            return Sheet.GetMergedRange(rowSpan, colSpan);
        }

        public IRowContainer NewRow(int rowSpan = 0)
        {
            return Sheet.NewRow(rowSpan);
        }

        public IRowContainer NewRow(XlsStyle style)
        {
            return Sheet.NewRow(style);
        }

        public IRowContainer NewRow(XlsStyle style, int rowSpan = 0)
        {
            return Sheet.NewRow(this, style, rowSpan);
        }

        public void SetColumn(int column = 1)
        {
            throw new NotImplementedException();
        }

        public IRangeContainer GetRoot()
        {
            return this;
        }

        public void Write(Stream stream)
        {
            ExcelPackage.SaveAs(stream);
        }
    }

    [Flags]
    public enum XlsFlags
    {
        Workbook = 1,
        Sheet = 2,
        Row = 4,
        Cell = 8,
        Container = 16,
        Text = 32,
        InsideSheet = Workbook | Sheet,
        InsideRow = InsideSheet | Row,
        InsideCell = InsideRow | Cell,
        TextContainer = InsideCell | Text | Container,
        CellContainer = InsideCell | Container,
        RowContainer = InsideRow | Container,
        SheetContainer = InsideSheet | Container
    }
}
