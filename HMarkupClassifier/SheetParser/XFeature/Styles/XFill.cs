using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    struct XFill
    {
        // 0-18 None is 17, solid is 18
        public int Pattern;
        public int PtnColor, BckColor;

        public XFill(IXLFill fill)
        {
            switch (fill.PatternType)
            {
                case XLFillPatternValues.None:
                    Pattern = 0;
                    break;
                case XLFillPatternValues.Solid:
                    Pattern = 1;
                    break;
                default:
                    Pattern = 2;
                    break;
            }
            PtnColor = fill.PatternColor.Color.ToArgb();
            BckColor = fill.BackgroundColor.Color.ToArgb();
        }

        public override int GetHashCode()
        {
            int hashCode = Pattern;
            hashCode = hashCode ^ PtnColor;
            hashCode = hashCode ^ BckColor;
            return hashCode;
        }

        public static string CSVTitle
            = "fill-pattern,fill-fcolor,fill-bcolor";

        public override string ToString()
        {
            return $"{Pattern},{PtnColor},{BckColor}";
        }
    }
}
