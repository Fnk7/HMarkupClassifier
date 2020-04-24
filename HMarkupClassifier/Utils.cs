using ClosedXML.Excel;


namespace HMarkupClassifier
{
    static class Utils
    {
        public static int ParseColumn(string col)
        {
            int temp = 0;
            foreach (var c in col)
                temp = temp * 26 + c - 'A' + 1;
            return temp;
        }

        public static int GetColor(XLColor xlColor)
        {
            switch (xlColor.ColorType)
            {
                case XLColorType.Theme:
                    int theme = (int)xlColor.ThemeColor;
                    int tint = (int)xlColor.ThemeTint * 10;
                    return (theme << 24) + tint;
                case XLColorType.Indexed:
                    return xlColor.Indexed;
                default:
                    return xlColor.Color.ToArgb();
            }
        }
    }
}
