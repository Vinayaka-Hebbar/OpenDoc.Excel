namespace OpenDoc.Excel
{
    public
#if LATEST_VS
        readonly
#endif
        struct PageBreak
    {
        public static readonly PageBreak None = new PageBreak();

        public readonly ExcelLimit[] Breaks;

        public PageBreak(ExcelLimit[] breaks)
        {
            Breaks = breaks;
        }

        public static PageBreak Parse(string value)
        {
            var split = value.Split(XmlHelpers.CommaSeperator);
            var breaks = new ExcelLimit[split.Length];
            if (split.Length > 0)
            {
                for (int index = 0; index < split.Length; index++)
                {
                    breaks[index] = ExcelLimit.Parse(split[index], XmlHelpers.DashSeparator);
                }
            }
            return new PageBreak(breaks);
        }
    }
}
