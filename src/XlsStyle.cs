using OfficeOpenXml.Style;
using System.Drawing;

namespace OpenDoc.Excel
{
    public struct XlsStyle : IDocStyle
    {
        internal static readonly XlsStyle None = new XlsStyle();
        internal ExcelHorizontalAlignment XlsHorizontalAlignment;
        internal ExcelVerticalAlignment XlsVerticalAlignment;
        Enums.HAlign HorizontalAlignment;
        Enums.VAlign VerticalAlignment;
        internal Enums.BorderType Border;
        internal ExcelBorderStyle BorderStyle;
        internal Color BorderColor;
        internal Color BgColor;
        internal DocFont Font;
        internal Enums.TextWrapType TextWrap;

        /// <summary>
        /// for table cols
        /// </summary>
        internal int TotalColumns;
        internal bool HasBorder => BorderStyle != ExcelBorderStyle.None;
        internal bool IsMerged => MergedRow | MergedCol;

        Color IDocStyle.BorderColor => BorderColor;

        Color IDocStyle.BgColor => BgColor;

        IDocFont IDocStyle.Font => Font;

        Enums.BorderType IDocStyle.Border => Border;

        DocSize IDocStyle.Size => Size;

        Enums.HAlign IDocStyle.HorizontalAlignment => HorizontalAlignment;

        Enums.VAlign IDocStyle.VerticalAlignment => VerticalAlignment;

        internal ExcelMerge MergedRow;
        internal ExcelMerge MergedCol;

        internal DocSize Size;

        internal XlsStyle(PageStyle style)
        {
            XlsVerticalAlignment = GetVAlign(style.VerticalAlignment);
            XlsHorizontalAlignment = GetHAlign(style.HorizontalAlignment);
            HorizontalAlignment = style.HorizontalAlignment;
            VerticalAlignment = style.VerticalAlignment;
            Border = Enums.BorderType.box;
            BorderStyle = ExcelBorderStyle.Thin;
            BgColor = Color.Empty;
            Font = style.Font;
            BorderColor = Color.LightGray;
            MergedRow = MergedCol = default(ExcelMerge);
            Size = DocSize.None;
            TotalColumns = -1;
            TextWrap = Enums.TextWrapType.nowrap;
        }

        internal XlsStyle(XlsStyle style)
        {
            XlsHorizontalAlignment = style.XlsHorizontalAlignment;
            XlsVerticalAlignment = style.XlsVerticalAlignment;
            HorizontalAlignment = style.HorizontalAlignment;
            VerticalAlignment = style.VerticalAlignment;
            Border = style.Border;
            BgColor = style.BgColor;
            BorderColor = style.BorderColor;
            MergedRow = style.MergedRow;
            MergedCol = style.MergedCol;
            Border = style.Border;
            BorderStyle = style.BorderStyle;
            Font = style.Font;
            Size = style.Size;
            TotalColumns = style.TotalColumns;
            TextWrap = style.TextWrap;
        }

        internal XlsStyle SetMergedCol(int start, int count)
        {
            MergedCol = MergedCol.Merge(start, count);
            return this;
        }

        internal void SetHAlign(Enums.HAlign alignment)
        {
            HorizontalAlignment = alignment;
            XlsHorizontalAlignment = GetHAlign(alignment);
        }

        internal void SetVAlign(Enums.VAlign alignment)
        {
            VerticalAlignment = alignment;
            XlsVerticalAlignment = GetVAlign(alignment);
        }

        internal void SetBorder(Enums.BorderType border)
        {
            Border = border;
        }

        internal XlsStyle SetMergedRow(int start, int skip)
        {
            MergedRow = MergedRow.Merge(start, skip);
            return this;
        }

        public override string ToString()
        {
            return $"Merged => {IsMerged}, Rows=> {MergedRow},Cols=> {MergedCol}";
        }

        private static ExcelHorizontalAlignment GetHAlign(Enums.HAlign alignment)
        {
            switch (alignment)
            {
                case Enums.HAlign.center:
                    return ExcelHorizontalAlignment.Center;
                case Enums.HAlign.left:
                    return ExcelHorizontalAlignment.Left;
                case Enums.HAlign.right:
                    return ExcelHorizontalAlignment.Right;
                case Enums.HAlign.justified:
                    return ExcelHorizontalAlignment.Justify;
                default:
                    return ExcelHorizontalAlignment.Fill;
            }
        }

        private static ExcelVerticalAlignment GetVAlign(Enums.VAlign alignment)
        {
            switch (alignment)
            {
                case Enums.VAlign.top:
                    return ExcelVerticalAlignment.Top;
                case Enums.VAlign.justified:
                    return ExcelVerticalAlignment.Justify;
                case Enums.VAlign.bottom:
                    return ExcelVerticalAlignment.Bottom;
                case Enums.VAlign.middle:
                case Enums.VAlign.baseline:
                default:
                    return ExcelVerticalAlignment.Center;
            }
        }
    }
}
