using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    struct XFont
    {
        public int Bold;
        public int Italic;
        public int Strikethrough;
        public int Underline;
        public int Vertical;

        public double Size;
        public int Color;
        public int NameIndex;

        public XFont(IXLFont font)
        {
            Bold = font.Bold ? 1 : 0;
            Italic = font.Italic ? 1 : 0;
            Strikethrough = font.Strikethrough ? 1 : 0;
            Underline = 0;
            switch (font.Underline)
            {
                case XLFontUnderlineValues.Double:
                case XLFontUnderlineValues.DoubleAccounting:
                    Underline = 2;
                    break;
                case XLFontUnderlineValues.Single:
                case XLFontUnderlineValues.SingleAccounting:
                    Underline = 1;
                    break;
            }
            Vertical = 0;
            switch (font.VerticalAlignment)
            {
                case XLFontVerticalTextAlignmentValues.Subscript:
                case XLFontVerticalTextAlignmentValues.Superscript:
                    Vertical = 1;
                    break;
            }
            Size = font.FontSize;
            Color = font.FontColor.Color.ToArgb();
            NameIndex = 0;
        }

        public override int GetHashCode()
        {
            int hashCode = Bold;
            hashCode = (hashCode << 2) ^ Italic;
            hashCode = (hashCode << 2) ^ Strikethrough;
            hashCode = (hashCode << 2) ^ Underline;
            hashCode = (hashCode << 2) ^ Vertical;
            hashCode = (hashCode << 2) ^ Color;
            hashCode = (hashCode << 2) ^ Size.GetHashCode();
            hashCode = (hashCode << 2) ^ NameIndex;
            return hashCode;
        }

        public static string CSVTitle
            = "fnt-bold,fnt-italic,fnt-strike,fnt-udrline,fnt-vertical,fnt-color,fnt-size,fnt-name";

        public override string ToString()
        {
            return $"{Bold},{Italic},{Strikethrough},{Underline},{Vertical},{Color},{Size},{NameIndex}";
        }
    }
}
