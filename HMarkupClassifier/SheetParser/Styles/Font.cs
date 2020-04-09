using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser.Styles
{
    class Font
    {
        public int Bold;
        public int Italic;
        public int Strikethrough;
        public int Underline;
        public int Vertical;

        public double Size;
        public int Color;
        public string Name;

        public int ColorIndex;
        public int NameIndex;

        public Font(IXLFont font)
        {
            Bold = font.Bold ? 1 : 0;
            Italic = font.Italic ? 1 : 0;
            Strikethrough = font.Strikethrough ? 1 : 0;
            switch (font.Underline)
            {
                case XLFontUnderlineValues.Double:
                case XLFontUnderlineValues.DoubleAccounting:
                    Underline = 2;
                    break;
                case XLFontUnderlineValues.None:
                    Underline = 0;
                    break;
                case XLFontUnderlineValues.Single:
                case XLFontUnderlineValues.SingleAccounting:
                    Underline = 1;
                    break;
            }
            switch (font.VerticalAlignment)
            {
                case XLFontVerticalTextAlignmentValues.Baseline:
                    Vertical = 0;
                    break;
                case XLFontVerticalTextAlignmentValues.Subscript:
                case XLFontVerticalTextAlignmentValues.Superscript:
                    Vertical = 1;
                    break;
            }
            Size = font.FontSize;
            Color = font.FontColor.Color.ToArgb();
            Name = font.FontName;
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
            hashCode = (hashCode << 2) ^ Name.GetHashCode();
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Font font = obj as Font;
            if (font == null) return false;
            if (Bold == font.Bold
                && Italic == font.Italic
                && Strikethrough == font.Strikethrough
                && Underline == font.Underline
                && Vertical == font.Vertical
                && Color == font.Color
                && Size == font.Size
                && Name == font.Name)
                return true;
            return false;
        }

        public static string CSVTitle
            = "fnt-bold,fnt-italic,fnt-strike,fnt-udrline,fnt-vertical,fnt-color,fnt-size,fnt-name";

        public override string ToString()
        {
            return $"{Bold},{Italic},{Strikethrough},{Underline},{Vertical},{ColorIndex},{Size},{NameIndex}";
        }
    }
}
