using ClosedXML.Excel;
using System;

namespace HMarkupClassifier.SheetParser
{
    struct XFill
    {
        // 0-18 None is 17, solid is 18
        public int Pattern;
        public int PtnColor, BckColor;

        public XFill(IXLFill fill)
        {
            Pattern = (int)fill.PatternType;
            PtnColor = Utils.GetColor(fill.PatternColor);
            BckColor = Utils.GetColor(fill.BackgroundColor);
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
