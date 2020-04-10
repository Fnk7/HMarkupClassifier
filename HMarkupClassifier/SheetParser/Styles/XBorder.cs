using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser.Styles
{
    struct XBorder
    {
        // 0-13 None is 10
        public int left, top, right, bottom;

        public XBorder(IXLBorder border)
        {
            left = GetBorderValue(border.LeftBorder);
            top = GetBorderValue(border.TopBorder);
            right = GetBorderValue(border.RightBorder);
            bottom = GetBorderValue(border.BottomBorder);
        }

        public static int GetBorderValue(XLBorderStyleValues value)
        {
            switch (value)
            {
                case XLBorderStyleValues.None:
                    return 0;
                case XLBorderStyleValues.Thin:
                    return 1;
                default:
                    return 2;
            }
        }

        public override int GetHashCode()
        {
            int hashCode = left;
            hashCode = (hashCode << 2) ^ top;
            hashCode = (hashCode << 2) ^ right;
            hashCode = (hashCode << 2) ^ bottom;
            return hashCode;
        }

        public static string CSVTitle
            = "bdr-left,bdr-top,bdr-right,bdr-bottom";

        public override string ToString()
        {
            return $"{left},{top},{right},{bottom}";
        }
    }
}
