using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenDoc.Internals;
using OpenDoc.Utils;
using OpenDoc.Xml;
using System;
using System.Drawing;
using System.IO;
using static OpenDoc.Tags;

namespace OpenDoc.Excel
{
    public sealed class ExcelTemplate : XmlTemplate<ExcelItem, IRangeContainer>
    {
        private const string COLUMNS = "columns";
        private const string INDEX = "index";

        public override XNamespace Namespace { get; } = XNamespace.Excel;

        private readonly ExcelItem None = ExcelItem.None;

        public ExcelTemplate(object model, XmlConfig config) : base(model, config)
        {
        }

        public override ITemplate<IRangeContainer> CreateDocument(XElement root)
        {
            return XlsWorkbook.Create(this, root);
        }

        protected override ExcelItem CreateSimpleTextElement(IRangeContainer parent, string value)
        {
            //todo
            var range = parent.GetRange();
            range.RichText.Add(value);
            return None;
        }

        protected override ExcelItem CreateEmptyRow(IRangeContainer parent, XElement element)
        {
            var rowCount = 1;
            if (element.HasAttributes)
            {
                var rows = element.Attribute(Namespace, ROWS);
                if (rows != null)
                {
                    rowCount = XmlConvert.ToInt32(rows);
                }
            }
            if ((parent.Flags & XlsFlags.Cell) == XlsFlags.Cell)
                parent.NewRow(rowCount);
            else
                parent.Sheet.MakeEmptyRows(rowCount);
            return None;
        }

        protected override ExcelItem EmptyElement(IRangeContainer parent)
        {
            parent.GetRange();
            return None;
        }


        protected override ExcelItem EmptySpace(IRangeContainer parent)
        {
            var range = parent.GetRange();
            range.RichText.Add(XmlHelpers.SpaceString);
            return None;
        }

        protected override ExcelItem CreateHeader(IRangeContainer parent, XElement element)
        {
            parent.Sheet.WorkSheet.HeaderFooter.FirstHeader.CenteredText = element.InnerText;
            return None;
        }

        protected override ExcelItem CreateFooter(IRangeContainer parent, XElement element)
        {
            parent.Sheet.WorkSheet.HeaderFooter.FirstFooter.CenteredText = element.InnerText;
            return None;
        }

        protected override ExcelItem CreateNewLine(IRangeContainer parent)
        {
            // if inside paragraph add newline
            if ((parent.Flags & XlsFlags.TextContainer) == 0)
                parent.NewRow();
            else
                parent.Sheet.MakeEmptyRow();
            return None;
        }

        protected override ExcelItem NextPage()
        {
            return None;
        }

        protected override ExcelItem CreateUnicodeElement(IRangeContainer parent, string value)
        {
            var range = parent.GetRange();
            range.Value = value;
            return None;
        }

        protected override ExcelItem CreateSimpleText(IRangeContainer parent, XElement element, int fontSize = 0, Enums.FontStyleType fontStyle = Enums.FontStyleType.inherited)
        {
            var style = parent.Style;
            ExcelRangeBase range = parent.GetRange();
            style.Font.Size += fontSize;
            if (fontStyle != Enums.FontStyleType.inherited)
                style.Font.Style = fontStyle;
            ExcelRichText text = range.RichText.Add(element.InnerText);
            text.SetFontStyle(style.Font);
            return None;
        }

        protected override ExcelItem CreateSimpleText(IRangeContainer parent, XElement element, DocFont font)
        {
            var style = parent.Style;
            ExcelRangeBase range = parent.GetRange();
            style.Font = style.Font.Set(font);
            ExcelRichText text = range.RichText.Add(element.InnerText);
            text.SetFontStyle(style.Font);
            return None;
        }

        protected override IRangeContainer CreateText(IRangeContainer parent, int fontSize = 0, Enums.FontStyleType fontStyle = Enums.FontStyleType.inherited)
        {
            var style = parent.Style;
            style.Font.Size += fontSize;
            if (fontStyle != Enums.FontStyleType.inherited)
                style.Font.Style = fontStyle;
            var range = parent.GetRange(style);
            return new TextContainer(parent.Sheet.CurrentRow, range, style);
        }

        protected override ExcelItem CreateLine()
        {
            return None;
        }

        protected override ExcelItem CreateTextElement(IRangeContainer parent, XElement element, int fontSize = 0)
        {
            var style = parent.Style;
            style.Font.Size += fontSize;
            ExcelRangeBase range = parent.GetRange(style);
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case COLOR:
                        style.Font.Color = GetColor(attr.Value);
                        break;
                    case FONTSIZE:
                        style.Font.Size = XmlConvert.ToInt32(attr);
                        break;
                    case FONTSTYLE:
                        Enum.TryParse(attr.Value, out style.Font.Style);
                        break;
                    case FONTFAMILY:
                        style.Font.Name = attr.Value;
                        range.Style.Font.SetFromFont(new Font(style.Font.Name, style.Font.Size));
                        //Todo
                        break;
                    case BORDER:
                        Enum.TryParse(attr.Value, out style.Border);
                        break;
                    case TEXTWRAP:
                        Enum.TryParse(attr.Value, out style.TextWrap);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                }
            }
            var text = range.RichText.Add(element.InnerText);
            //Apply style if not a row
            if (style.TextWrap != Enums.TextWrapType.nowrap)
                range.Style.WrapText = true;
            if ((parent.Flags & XlsFlags.Cell) == XlsFlags.Cell)
            {
                text.SetFontStyle(style.Font);
            }
            return None;
        }

        protected override IRangeContainer Default(IRangeContainer parent, XElement element)
        {
            switch (element.Name)
            {
                case COLUMNS:
                    SetColumnsWidth(parent, element);
                    return RowContainer.Empty;
            }
            return new ElementContainer(parent, false, false);
        }

        private void SetColumnsWidth(IRangeContainer parent, XElement element)
        {
            foreach (var item in element.Elements(Namespace))
            {
                SetColumnWidth(parent, item);
            }
        }

        private void SetColumnWidth(IRangeContainer parent, XElement item)
        {
            int index = -1;
            double minWidth = 0;
            double maxWidth = 0;
            double width = -1;
            foreach (var attr in item.Attributes())
            {
                switch (attr.Name)
                {
                    case INDEX:
                        index = XmlConvert.ToInt32(attr);
                        break;
                    case MINWIDTH:
                        minWidth = XmlConvert.ToDouble(attr);
                        break;
                    case MAXWIDTH:
                        maxWidth = XmlConvert.ToDouble(attr);
                        break;
                    case WIDTH:
                        if (Enum.TryParse(attr.Value, out Enums.SizeType widthType))
                        {
                            width = (sbyte)widthType;
                            break;
                        }
                        width = XmlConvert.ToDouble(attr);
                        break;
                }
            }
            if (index != -1)
            {
                var sheet = parent.Sheet.WorkSheet;
                var column = sheet.Column(index);
                column.AutoFit(minWidth, maxWidth);
                if (!double.IsNaN(width))
                {
                    switch (width)
                    {
                        //Auto
                        case -1:
                            column.AutoFit();
                            break;
                        //Fit
                        case -2:
                            column.BestFit = true;
                            break;
                        default:
                            column.Width = width;
                            break;
                    }

                }
            }
        }

        protected override ExcelItem CreateImage(IRangeContainer parent, XElement element)
        {
            SetDataContext(element);
            int height = 800;
            int width = 600;
            int row = parent.Row;
            int column = parent.Column;
            //todo image http source
            string src = null;
            string name = string.Empty;
            bool shouldDispose = false;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case NAME:
                        name = attr.Value;
                        break;
                    case HEIGHT:
                        height = XmlConvert.ToInt32(attr);
                        break;
                    case WIDTH:
                        width = XmlConvert.ToInt32(attr);
                        break;
                    case ROW:
                        row = XmlConvert.ToInt32(attr);
                        break;
                    case COLUMN:
                        row = XmlConvert.ToInt32(attr);
                        break;
                    case SRC:
                        src = attr.Value;
                        break;
                    case PATH:
                        shouldDispose = true;
                        element.DataContext = GetStreamObject(attr.Value);
                        break;
                }
            }
            Image image = null;
            if (element.DataContext is Stream stream)
            {
                image = Image.FromStream(stream);
                if (shouldDispose)
                    stream.Dispose();
            }
            if (image != null)
            {
                var picture = parent.Sheet.WorkSheet.Drawings.AddPicture(name, image);
                picture.SetPosition(row, 0, column, 0);
                picture.SetSize(width, height);
            }
            return None;
        }

        protected override ExcelItem CreateDrawing(IRangeContainer parent, XElement element)
        {
            var sheet = parent.Sheet.WorkSheet;
            int width = 50, height = 50;
            string name = null;
            var shapeType = eShapeStyle.Line;
            var fillColor = Color.White;
            var borderColor = Color.Black;
            var content = string.Empty;
            int row = 0, column = 0;
            double rowHeigth = 0;
            OfficeOpenXml.Drawing.eTextAnchoringType anchoringType = OfficeOpenXml.Drawing.eTextAnchoringType.Center;
            int offsetX = 0, offsetY = 0;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case NAME:
                        name = attr.Value;
                        break;
                    case "row-height":
                        rowHeigth = XmlConvert.ToDouble(attr);
                        break;
                    case HEIGHT:
                        height = XmlConvert.ToInt32(attr);
                        break;
                    case WIDTH:
                        width = XmlConvert.ToInt32(attr);
                        break;
                    case "shape":
                        shapeType = (eShapeStyle)Enum.Parse(typeof(eShapeStyle), attr.Value);
                        break;
                    case "fill-color":
                        fillColor = GetColor(attr.Value);
                        break;
                    case "border-color":
                        borderColor = GetColor(attr.Value);
                        break;
                    case CONTENT:
                        content = attr.Value;
                        break;
                    case ROW:
                        row = XmlConvert.ToInt32(attr);
                        break;
                    case COLUMN:
                        column = XmlConvert.ToInt32(attr);
                        break;
                    case "offset-x":
                        offsetX = XmlConvert.ToInt32(attr);
                        break;
                    case "offset-y":
                        offsetY = XmlConvert.ToInt32(attr);
                        break;
                    case "justify":
                        anchoringType = (OfficeOpenXml.Drawing.eTextAnchoringType)Enum.Parse(typeof(OfficeOpenXml.Drawing.eTextAnchoringType), attr.Value);
                        break;

                }
            }
            var shape = sheet.Drawings.AddShape(name, shapeType);
            shape.Fill.Color = fillColor;
            shape.Border.Fill.Color = borderColor;
            if (row > 0 && rowHeigth > 0)
            {
                sheet.Row(row).Height = rowHeigth;
            }
            shape.SetSize(width, height);
            //offset Y row, offsetX column dir
            shape.SetPosition(row, offsetY, column, offsetX);
            shape.TextAnchoring = anchoringType;
            shape.Text = content;
            shape.Font.Apply(parent.Style);
            return ExcelItem.None;
        }

        protected override IRangeContainer CreateChart(IRangeContainer parent, XElement element)
        {
            ChartConfig config = new ChartConfig()
            {
                Width = 800,
                Height = 600,
                LegendEnabled = true,
                LegendPosition = 2,
                XFormat = null,
                YFormat = null
            };
            var chartType = OfficeOpenXml.Drawing.Chart.eChartType.Line;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case HEIGHT:
                        config.Height = XmlConvert.ToInt32(attr);
                        break;
                    case WIDTH:
                        config.Width = XmlConvert.ToInt32(attr);
                        break;
                    case "x-label":
                        config.XLabel = attr.Value;
                        break;
                    case "y-label":
                        config.YLabel = attr.Value;
                        break;
                    case "title":
                        config.Title = attr.Value;
                        break;
                    case "legend-enabled":
                        config.LegendEnabled = XmlConvert.ToBoolean(attr);
                        break;
                    case "legend-position":
                        config.LegendPosition = (int)Enum.Parse(typeof(OfficeOpenXml.Drawing.Chart.eLegendPosition), attr.Value);
                        break;
                    case "x-limit":
                        config.XAxisLimit = AxisLimit.Parse(attr.Value);
                        break;
                    case "y-limit":
                        config.YAxisLimit = AxisLimit.Parse(attr.Value);
                        break;
                    case "unit-limit":
                        config.Units = AxisLimit.Parse(attr.Value);
                        break;
                    case "x-format":
                        config.XFormat = attr.Value;
                        break;
                    case "y-format":
                        config.YFormat = attr.Value;
                        break;
                    case TYPE:
                        chartType = (OfficeOpenXml.Drawing.Chart.eChartType)Enum.Parse(typeof(OfficeOpenXml.Drawing.Chart.eChartType), attr.Value);
                        break;
                    case ROW:
                        config.Row = XmlConvert.ToInt32(attr);
                        break;
                    case COLUMN:
                        config.Column = XmlConvert.ToInt32(attr);
                        break;
                    case "rounded-corner":
                        config.RoundedCorners = XmlConvert.ToBoolean(attr);
                        break;
                    case "border":
                        config.Border.Style = (int)Enum.Parse(typeof(OfficeOpenXml.Drawing.Chart.eChartType), attr.Value);
                        break;

                }
            }
            if (element.HasElements)
            {
                var excelChart = new ExcelChart(this);
                OfficeOpenXml.Drawing.Chart.ExcelChart chart = excelChart.Create(parent, element, chartType, config.Title);
                if (config.XAxisLimit != null && config.XAxisLimit.HasValue)
                {
                    var xAxisLimit = config.XAxisLimit.Value;
                    chart.XAxis.MaxValue = xAxisLimit.Max;
                    chart.XAxis.MinValue = xAxisLimit.Min;
                }
                if (config.YAxisLimit != null && config.YAxisLimit.HasValue)
                {
                    var yAxisLimit = config.YAxisLimit.Value;
                    chart.YAxis.MaxValue = yAxisLimit.Max;
                    chart.YAxis.MinValue = yAxisLimit.Min;
                }
                chart.SetSize(config.Width, config.Height);
                chart.SetPosition(config.Row, 0, config.Column, 0);
                chart.Title.Text = config.Title;
                chart.RoundedCorners = config.RoundedCorners;
                chart.XAxis.Title.Text = config.XLabel;
                chart.YAxis.Title.Text = config.YLabel;
                if (config.XFormat != null)
                {
                    chart.XAxis.Format = config.XFormat;
                }
                if (config.YFormat != null)
                {
                    chart.YAxis.Format = config.YFormat;
                }
                if (config.LegendEnabled)
                {
                    chart.Legend.Position = (OfficeOpenXml.Drawing.Chart.eLegendPosition)config.LegendPosition;
                }
                else
                {
                    chart.Legend.Remove();
                }
            }
            return new TemplaeContainer(parent, false, false);
        }

        protected override IRangeContainer CreateTemplateModel(IRangeContainer parent, XElement element)
        {
            var attr = element.Attribute(NAME);
            element.DataContext = _model.TryGetValue(attr.Value);
            return new TemplaeContainer(parent, true, false);
        }

        protected override IRangeContainer CreateModel(IRangeContainer parent, XElement element)
        {
            SetDataContext(element);
            if (element.HasElements)
                return parent.NewRow();
            var range = parent.GetRange();
            range.Value = element.InnerText;
            return new TemplaeContainer(parent, true, false);
        }

        protected override IRangeContainer GetTemplateContainer(IRangeContainer parent)
        {
            return new TemplaeContainer(parent, true, false);
        }

        protected override IRangeContainer CreateTableRow(IRangeContainer parent, XElement element)
        {
            XlsStyle style = parent.Style;
            SetDataContext(element);
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case HEIGHT:
                        style.Size.Height = XmlConvert.ToSingle(attr);
                        break;
                    case BORDER:
                        Enum.TryParse(attr.Value, out style.Border);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                    case FONTSIZE:
                        style.Font.Size = XmlConvert.ToSingle(attr);
                        break;
                    case FONTSTYLE:
                        Enum.TryParse(attr.Value, out style.Font.Style);
                        break;
                }
            }
            return parent.NewRow(style).ApplyStyle();
        }

        protected override IRangeContainer CreateTableCell(IRangeContainer parent, XElement element)
        {
            int colSpan = 1;
            int rowSpan = 1;
            int column = parent.Column;
            XlsStyle style = parent.Style;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case WIDTH:
                        string value = attr.Value;
                        if (Enum.TryParse(value, out Enums.SizeType widthType))
                        {
                            style.Size.Width = (sbyte)widthType;
                            break;
                        }
                        style.Size.Width = float.Parse(value);
                        break;
                    case BORDER:
                        Enum.TryParse(attr.Value, out Enums.BorderType border);
                        style.Border = border;
                        break;
                    case COLSPAN:
                        colSpan = XmlConvert.ToInt32(attr);
                        style = style.SetMergedCol(column, colSpan);
                        break;
                    case ROWSPAN:
                        rowSpan = XmlConvert.ToInt32(attr);
                        style = style.SetMergedRow(column, rowSpan);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                }
            }
            var range = parent.GetMergedRange(rowSpan, colSpan);
            if (!element.HasElements)
            {
                // to check whether value is formated
                range.Value = element.FirstNode is XText t ? t.FormattedContent() : element.Value;
            }
            if (!float.IsNaN(style.Size.Width))
            {
                var width = style.Size.Width;
                switch (width)
                {
                    //Auto
                    case -1:
                        range.AutoFitColumns();
                        break;
                    //Fit
                    case -2:
                        range.Style.ShrinkToFit = true;
                        break;
                    default:
                        var sheet = parent.Sheet;
                        var col = sheet.WorkSheet.Column(column);
                        col.Width = width;
                        break;
                }

            }
            range.SetStyle(style);
            return new CellContainer(parent.Sheet.CurrentRow, range, style);
        }

        protected override IRangeContainer CreateTableHeaderCell(IRangeContainer parent, XElement element)
        {
            return CreateTableCell(parent, element);
        }

        protected override IRangeContainer CreateTable(IRangeContainer parent, XElement element)
        {
            var style = parent.Style;
            var container = parent;
            int column = parent.Column;
            int row = parent.Row;
            SetDataContext(element);
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case BORDER:
                        Enum.TryParse(attr.Value, out Enums.BorderType border);
                        style.Border = border;
                        break;
                    case BORDERCOLOR:
                        style.BorderColor = GetColor(attr.Value);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                    case ROWHEIGHT:
                        style.Size.Height = XmlConvert.ToSingle(attr);
                        break;
                    case COLOR:
                        style.Font.Color = GetColor(attr.Value);
                        break;
                    case FONTSIZE:
                        style.Font.Size = XmlConvert.ToInt32(attr);
                        break;
                    case FONTSTYLE:
                        Enum.TryParse(attr.Value, out Enums.FontStyleType fontStyle);
                        style.Font.Style = fontStyle;
                        break;
                    case FONTFAMILY:
                        //Todo
                        break;
                    case FONTNAME:
                        //Todo
                        break;
                    case ROW:
                        container = new ExcelSheetWrapper(parent, XmlConvert.ToInt32(attr), style);
                        break;
                    case COLUMN:
                        column = XmlConvert.ToInt32(attr);
                        break;
                    case COLS:
                        style.TotalColumns = XmlConvert.ToInt32(attr);
                        break;
                }
            }
            return new ElementContainer(container, row, column, style);
        }

        protected override IRangeContainer CreateListItem(IRangeContainer parent, XElement element)
        {
            var row = parent.NewRow();
            var range = row.GetRange();
            if (!element.HasElements)
            {
                var text = range.RichText.Add(element.InnerText);
                if ((parent.Flags & XlsFlags.Cell) == XlsFlags.Cell)
                    text.SetFontStyle(parent.Style.Font);
                return row;
            }
            return new CellContainer(row, range);
        }

        protected override IRangeContainer CreateBorder(IRangeContainer parent, XElement element)
        {
            var style = parent.Style;
            SetDataContext(element);
            ISheetContainer sheet = parent.Sheet;
            IRowContainer currentRow = sheet.CurrentRow;
            int startColumn, column;
            int startRow, row;
            row = startRow = currentRow.Row;
            startColumn = column = currentRow.Column;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case BORDER:
                        Enum.TryParse(attr.Value, out Enums.BorderType border);
                        style.Border = border;
                        break;
                    case BORDERCOLOR:
                        style.BorderColor = GetColor(attr.Value);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                    case COLS:
                        column = XmlConvert.ToInt32(attr);
                        break;
                    case ROWS:
                        row = XmlConvert.ToInt32(attr);
                        break;
                    case ROW:
                        startRow = XmlConvert.ToInt32(attr);
                        parent = new ExcelSheetWrapper(parent, startRow, style);
                        break;
                    case COLUMN:
                        startColumn = XmlConvert.ToInt32(attr);
                        parent.SetColumn(startColumn);
                        break;
                }
            }
            var container = new ElementContainer(parent, false, false);
            //Iterate From Here
            IterateNodes(container, element.Childs);
            element.RemoveNodes();
            var range = sheet.WorkSheet.Cells[startRow, startColumn, startRow + row, startColumn + column];
            range.SetStyle(style);
            return container;
        }

        protected override IRangeContainer CreateDiv(IRangeContainer parent, XElement element)
        {
            var container = parent;
            var style = parent.Style;
            int column = parent.Column;
            int row = parent.Row;
            int rows = 1;
            int cols = 1;
            SetDataContext(element);
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case BORDER:
                        Enum.TryParse(attr.Value, out Enums.BorderType border);
                        style.Border = border;
                        break;
                    case BORDERCOLOR:
                        style.BorderColor = GetColor(attr.Value);
                        break;
                    case HEIGHT:
                        style.Size.Height = XmlConvert.ToSingle(attr);
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                    case ROW:
                        row = XmlConvert.ToInt32(attr);
                        container = new ExcelSheetWrapper(parent, row, style);
                        break;
                    case COLUMN:
                        column = XmlConvert.ToInt32(attr);
                        break;
                    case COLS:
                        cols = XmlConvert.ToInt32(attr);
                        break;
                    case ROWS:
                        rows = XmlConvert.ToInt32(attr);
                        break;
                }
            }
            container = new ElementContainer(container, row, column, style, false);
            ExcelRangeBase range;
            range = container.Sheet.WorkSheet.Cells[row, column, row + rows - 1, column + cols - 1];
            ExcelHelper.SetBorderStyle(range.Style.Border, style);
            //Iterate From Here
            IterateNodes(container, element.Childs);
            return container;
        }

        protected override IRangeContainer CreatePage(IRangeContainer parent, XElement element)
        {
            var template = parent.Workbook;
            return template.Create(element);
        }

        protected override IRangeContainer CreateText(IRangeContainer parent, XElement element, int fontSize = 0)
        {
            var style = parent.Style;
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case COLOR:
                        style.Font.Color = GetColor(attr.Value);
                        break;
                    case FONTSIZE:
                        style.Font.Size = XmlConvert.ToInt32(attr);
                        break;
                    case FONTSTYLE:
                        Enum.TryParse(attr.Value, out Enums.FontStyleType fontStyle);
                        style.Font.Style = fontStyle;
                        break;
                    case BGCOLOR:
                        style.BgColor = GetColor(attr.Value);
                        break;
                    case FONTFAMILY:
                        //Todo
                        break;
                    case FONTNAME:
                        style.Font.Name = attr.Value;
                        //Todo
                        break;
                    case TEXTWRAP:
                        Enum.TryParse(attr.Value, out style.TextWrap);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;

                }
            }

            var range = parent.GetRange(style);
            return new TextContainer(parent.Sheet.CurrentRow, range, style);
        }

        protected override ExcelItem CreateTextElement(IRangeContainer parent, string text)
        {
            var range = parent.GetRange();
            range.RichText.Add(text);
            return None;
        }

        protected override ExcelItem CreateNumber(IRangeContainer parent, XElement element)
        {
            //todo style applied?
            var number = GetNumber(element);
            var range = parent.GetRange();
            if ((parent.Flags & XlsFlags.TextContainer) == XlsFlags.TextContainer)
                range.RichText.Add(number.ToString());
            else
                range.Value = number;
            return None;
        }

        protected override IRangeContainer CreateList(IRangeContainer parent, XElement element, bool isOrdered)
        {
            SetDataContext(element);
            return new TemplaeContainer(parent, true, false);
        }

        protected override IRangeContainer CreateParagraph(IRangeContainer parent, XElement element)
        {
            var style = parent.Style;
            int colspan = 1;
            var template = parent.NewRow();
            ExcelRow excelRow = parent.Sheet.WorkSheet.Row(template.Row);
            SetDataContext(element);
            foreach (var attr in element.Attributes(Namespace))
            {
                switch (attr.Name)
                {
                    case COLSPAN:
                        colspan = XmlConvert.ToInt32(attr);
                        style = style.SetMergedCol(template.Column, colspan);
                        break;
                    case HEIGHT:
                        excelRow.Height = XmlConvert.ToDouble(attr);
                        break;
                    case HALIGN:
                        Enum.TryParse(attr.Value, out Enums.HAlign align);
                        style.SetHAlign(align);
                        break;
                    case VALIGN:
                        Enum.TryParse(attr.Value, out Enums.VAlign valign);
                        style.SetVAlign(valign);
                        break;
                }
            }
            var range = template.GetRange(style);
            range.Style.WrapText = true;
            //Iterate From Here
            IterateNodes(new TextContainer(template, range, style), element.Childs);
            if (range.IsRichText)
            {
                var sheet = parent.Sheet;
                int count = range.RichText.Count;
                template.MergedRange(template.Row, 1, count, colspan);
                sheet.MakeEmptyRows(count);
            }
            return null;
        }
    }


#region Excel Style
    internal struct ExcelMerge
    {
        internal int Count;
        internal int Start;
        internal bool IsMerged;

        internal int Total
        {
            get
            {
                return Start + Count;
            }
        }

        public ExcelMerge(int start, int skip, bool isMerged)
        {
            Start = start;
            Count = skip;
            IsMerged = isMerged;
        }

        public bool IsBetween(int value)
        {
            return Start < value && value <= Total;
        }

        public static int operator -(ExcelMerge merge, int number)
        {
            return merge.Count - number;
        }

        public static bool operator |(ExcelMerge merge1, ExcelMerge merge2)
        {
            return merge1.IsMerged || merge2.IsMerged;
        }

        public ExcelMerge Merge(int start, int number)
        {
            return new ExcelMerge(start, number, number > 1);
        }
    }
#endregion

    internal static class ExcelHelper
    {
        internal static void SetStyle(this ExcelRangeBase range, XlsStyle style)
        {
            ExcelStyle xlsStyle = range.Style;
            Apply(range.Style, style);
            if (style.HasBorder)
                SetBorderStyle(xlsStyle.Border, style);
        }

        internal static void Apply(ExcelStyle xlsStyle, XlsStyle style)
        {
            xlsStyle.HorizontalAlignment = style.XlsHorizontalAlignment;
            xlsStyle.VerticalAlignment = style.XlsVerticalAlignment;
            if (style.BgColor != Color.Empty)
            {
                xlsStyle.Fill.PatternType = ExcelFillStyle.Solid;
                xlsStyle.Fill.BackgroundColor.SetColor(style.BgColor);
            }
            if (style.TextWrap != Enums.TextWrapType.nowrap)
                xlsStyle.WrapText = true;
            SetFont(xlsStyle.Font, style);
        }

        internal static void SetStyle(this ExcelRow row, XlsStyle style)
        {
            //Can't set background and border here
            var rangeStyle = row.Style;
            //can't apply alignment here due to entire row alignment will change
            SetFont(rangeStyle.Font, style);
            if (!float.IsNaN(style.Size.Height))
            {
                row.Height = style.Size.Height;
            }
        }

        internal static void SetBorderStyle(OfficeOpenXml.Style.Border border, XlsStyle style)
        {
            switch (style.Border)
            {
                case Enums.BorderType.none:
                    break;
                case Enums.BorderType.top:
                    border.Top.Style = style.BorderStyle;
                    border.Top.Color.SetColor(style.BorderColor);
                    break;
                case Enums.BorderType.bottom:
                    border.Bottom.Style = style.BorderStyle;
                    border.Bottom.Color.SetColor(style.BorderColor);
                    break;
                case Enums.BorderType.left:
                    border.Left.Style = style.BorderStyle;
                    border.Left.Color.SetColor(style.BorderColor);
                    break;
                case Enums.BorderType.right:
                    border.Right.Style = style.BorderStyle;
                    border.Right.Color.SetColor(style.BorderColor);
                    break;
                case Enums.BorderType.box:
                    border.BorderAround(style.BorderStyle, style.BorderColor);
                    break;
            }
        }

        internal static void SetFontStyle(this ExcelRichText richText, DocFont font)
        {
            richText.Color = font.Color;
            richText.Size = font.Size;
            richText.FontName = font.Name;
            SetFontStyle(richText, font.Style);
            switch (font.VAlign)
            {
                case Enums.VAlign.sub:
                    richText.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
                    break;
                case Enums.VAlign.super:
                    richText.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
                    break;
            }
        }

        internal static void Apply(this ExcelTextFont xlsFont, XlsStyle style)
        {
            var font = style.Font;
            xlsFont.Color = font.Color;
            xlsFont.Size = font.Size;
            switch (font.Style)
            {
                case Enums.FontStyleType.inherited:
                    return;
                case Enums.FontStyleType.normal:
                    xlsFont.Bold = false;
                    xlsFont.Italic = false;
                    return;
                case Enums.FontStyleType.bold:
                    xlsFont.Bold = true;
                    return;
                case Enums.FontStyleType.italic:
                    xlsFont.Italic = true;
                    break;
                case Enums.FontStyleType.bolditalic:
                    xlsFont.Bold = true;
                    xlsFont.Italic = true;
                    break;
                default:
                    break;
            }
        }

        internal static void SetFont(ExcelFont xlsFont, XlsStyle style)
        {
            //todo avoid re assign of styles
            var font = style.Font;
            xlsFont.Color.SetColor(font.Color);
            xlsFont.Size = font.Size;
            xlsFont.Name = style.Font.Name;
            switch (font.Style)
            {
                case Enums.FontStyleType.inherited:
                    return;
                case Enums.FontStyleType.normal:
                    xlsFont.Bold = false;
                    xlsFont.Italic = false;
                    xlsFont.UnderLine = false;
                    return;
                case Enums.FontStyleType.bold:
                    xlsFont.Bold = true;
                    return;
                case Enums.FontStyleType.italic:
                    xlsFont.Italic = true;
                    break;
                case Enums.FontStyleType.bolditalic:
                    xlsFont.Bold = true;
                    xlsFont.Italic = true;
                    break;
                case Enums.FontStyleType.underline:
                    xlsFont.UnderLineType = ExcelUnderLineType.Single;
                    xlsFont.UnderLine = true;
                    break;
                default:
                    break;
            }
        }

        internal static void SetFontStyle(ExcelRichText text, Enums.FontStyleType style)
        {
            switch (style)
            {
                case Enums.FontStyleType.inherited:
                    return;
                case Enums.FontStyleType.normal:
                    text.Bold = false;
                    text.Italic = false;
                    text.UnderLine = false;
                    return;
                case Enums.FontStyleType.bold:
                    text.Bold = true;
                    return;
                case Enums.FontStyleType.italic:
                    text.Italic = true;
                    break;
                case Enums.FontStyleType.bolditalic:
                    text.Bold = true;
                    text.Italic = true;
                    break;
                case Enums.FontStyleType.underline:
                    text.UnderLine = true;
                    break;
                default:
                    break;
            }
        }
    }
}
