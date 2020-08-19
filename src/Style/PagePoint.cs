namespace OpenDoc.Excel
{
    internal
#if LATEST_VS
        readonly
#endif
        struct PagePoint
    {

        internal static PagePoint Zero = new PagePoint();

        public readonly int Left;
        public readonly int Top;
        public readonly int Right;
        public readonly int Bottom;

        internal PagePoint(int left, int top, int right, int bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal bool IsGreaterThan(int value)
        {
            return Left > value && Top > value && Right > value && Bottom > value
                && Right > Left && Bottom > Top;
        }

        internal static PagePoint Parse(string value, char separator = XmlHelpers.CommaSeperator)
        {
            var values = value.Split(separator);
            int left, top, right, bottom;
            switch (values.Length)
            {
                case 4:
                    left = int.Parse(values[0]);
                    top = int.Parse(values[1]);
                    right = int.Parse(values[2]);
                    bottom = int.Parse(values[3]);
                    break;
                case 2:
                    left = right = int.Parse(values[0]);
                    top = bottom = int.Parse(values[1]);
                    break;
                case 1:
                    left = right = top = bottom = int.Parse(values[0]);
                    break;
                default:
                    left = right = top = bottom = 0;
                    break;
            }
            return new PagePoint(left, top, right, bottom);
        }
    }
}
