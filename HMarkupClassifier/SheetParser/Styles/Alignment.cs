using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser.Styles
{
    class Alignment
    {
        // TODO Alignment
        // 0-7 General is 4
        int Horizontal;
        // 0-4 Bottom is 0
        int Vertical;
        int Indent;
        int WrapText;
        int ShrinkToFit;

        public Alignment(IXLAlignment alignment)
        {
            Horizontal = (int)alignment.Horizontal;
            Vertical = (int)alignment.Vertical;
            Indent = alignment.Indent == 0 ? 0 : 1;
            WrapText = alignment.WrapText ? 1 : 0;
            ShrinkToFit = alignment.ShrinkToFit ? 1 : 0;
        }

        public override int GetHashCode()
        {
            int hashCode = Horizontal;
            hashCode = hashCode * 127 + Vertical;
            hashCode = hashCode * 127 + Indent;
            hashCode = hashCode * 127 + WrapText;
            hashCode = hashCode * 127 + ShrinkToFit;
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Alignment alignment = obj as Alignment;
            if (alignment == null) return false;
            if (Horizontal == alignment.Horizontal
                && Vertical == alignment.Vertical
                && Indent == alignment.Indent
                && WrapText == alignment.WrapText
                && ShrinkToFit == alignment.ShrinkToFit)
                return true;
            return false;
        }

        public static string CSVTitle
            = "alg-horizontal,alg-vertical,alg-indent,alg-wrap,alg-shrink";

        public override string ToString()
        {
            return $"{Horizontal},{Vertical},{Indent},{WrapText},{ShrinkToFit}";
        }
    }
}
