using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    struct XAlignment
    {
        // 0-7 General is 4
        public int Horizontal;
        // 0-4 Bottom is 0
        public int Vertical;
        public int Indent;
        public int TextRotation;
        public int WrapText;
        public int ShrinkToFit;

        public XAlignment(IXLAlignment alignment)
        {
            
            Horizontal = (int)alignment.Horizontal;
            Vertical = (int)alignment.Vertical;
            Indent = alignment.Indent;
            TextRotation = alignment.TextRotation;
            WrapText = alignment.WrapText ? 1 : 0;
            ShrinkToFit = alignment.ShrinkToFit ? 1 : 0;
        }

        public override int GetHashCode()
        {
            int hashCode = Horizontal;
            hashCode = (hashCode << 2) ^ Vertical;
            hashCode = (hashCode << 2) ^ Indent;
            hashCode = (hashCode << 2) ^ TextRotation;
            hashCode = (hashCode << 2) ^ WrapText;
            hashCode = (hashCode << 2) ^ ShrinkToFit;
            return hashCode;
        }

        public static string CSVTitle
            = "alg-horizontal,alg-vertical,alg-indent,alg-rotation,alg-wrap,alg-shrink";

        public override string ToString()
        {
            return $"{Horizontal},{Vertical},{Indent},{TextRotation},{WrapText},{ShrinkToFit}";
        }
    }
}
