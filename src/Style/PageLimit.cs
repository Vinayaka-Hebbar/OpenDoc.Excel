namespace OpenDoc.Excel
{
    public
#if LATEST_VS
    readonly
#endif
        struct ExcelLimit
    {
        public static readonly ExcelLimit Zero = new ExcelLimit();
        public readonly int Row;
        public readonly int Col;

        internal ExcelLimit(int row, int col)
        {
            Row = row;
            Col = col;
        }

        internal bool IsGreaterThan(int limit)
        {
            return Row > limit || Col > limit;
        }

        internal static ExcelLimit Parse(string value, char separator = XmlHelpers.CommaSeperator)
        {
            var split = value.Split(separator);
            if (split.Length == 1)
            {
                int.TryParse(value, out int row);
                return new ExcelLimit(row, 0);
            }
            if (split.Length == 2)
            {
                int.TryParse(split[0], out int row);
                int.TryParse(split[1], out int col);
                return new ExcelLimit(row, col);
            }
            return Zero;
        }
    }
}
