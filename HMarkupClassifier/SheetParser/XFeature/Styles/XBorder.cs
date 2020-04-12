using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    struct XBorder
    {
        // 0-13 None is 10
        public int left, top, right, bottom;

        public XBorder(IXLBorder border)
        {
            left = (int)border.LeftBorder;
            top = (int)border.TopBorder;
            right = (int)border.RightBorder;
            bottom = (int)border.BottomBorder;
        }

        public override int GetHashCode()
        {
            int hashCode = left;
            hashCode = (hashCode << 5) ^ top;
            hashCode = (hashCode << 5) ^ right;
            hashCode = (hashCode << 5) ^ bottom;
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
