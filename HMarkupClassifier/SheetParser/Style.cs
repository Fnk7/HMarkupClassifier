using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class Style
    {
        public int count = 0;

        public int prefix;
        public int numFormat;
        public Alignment alg;
        public Border bdr;
        public Fill fill;
        public Font font;

        public string CSV = null;

        public Style(IXLStyle style)
        {
            prefix = (byte)(style.IncludeQuotePrefix ? 1 : 0);
            numFormat = (byte)(style.NumberFormat.NumberFormatId == 0 ? 0 : 1);
            alg = new Alignment(style.Alignment);
            bdr = new Border(style.Border);
            fill = new Fill(style.Fill);
            font = new Font(style.Font);
        }

        public static string CSVTitle
            = $"style-prefix,style-numformat,{Alignment.CSVTitle},{Border.CSVTitle},{Fill.CSVTitle},{Font.CSVTitle}";


        public string ToCSV()
        {
            if (CSV == null)
                CSV = $"{prefix},{numFormat},{alg.ToCSV()},{bdr.ToCSV()},{fill.ToCSV()},{font.ToCSV()}";
            return CSV;
        }

        public override int GetHashCode()
        {
            int hashCode = prefix;
            hashCode = hashCode * 127 + numFormat;
            hashCode = hashCode * 127 + alg.GetHashCode();
            hashCode = hashCode * 127 + bdr.GetHashCode();
            hashCode = hashCode * 127 + fill.GetHashCode();
            hashCode = hashCode * 127 + font.GetHashCode();
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Style style = obj as Style;
            if (obj == null) return false;
            if (prefix == style.prefix
                && numFormat == style.numFormat
                && alg.Equals(style.alg)
                && bdr.Equals(style.bdr)
                && fill.Equals(style.fill)
                && font.Equals(style.font))
                return true;
            return false;
        }

        public string CSVString()
        {
            if (CSV == null)
                CSV = $"{prefix},{numFormat},{alg.CSVString()},{bdr.CSVString()},{fill.CSVString()},{font.CSVString()}";
            return CSV;
        }
    }

    class Alignment
    {
        // 0-7 General is 4
        int horizontal;
        // 0-4 Bottom is 0
        int vertical;
        int indent;
        int wrapText;
        int shrinkToFit;

        public Alignment(IXLAlignment alignment)
        {
            horizontal = (int)alignment.Horizontal;
            vertical = (int)alignment.Vertical;
            indent = alignment.Indent == 0 ? 0 : 1;
            wrapText = alignment.WrapText ? 1 : 0;
            shrinkToFit = alignment.ShrinkToFit ? 1 : 0;
        }

        public static string CSVTitle
            = "alg-hor,alg-ver,alg-indent,alg-wrap,alg-shrink";

        public string ToCSV()
            => $"{horizontal},{vertical},{indent},{wrapText},{shrinkToFit}";

        public override int GetHashCode()
        {
            int hashCode = horizontal;
            hashCode = hashCode * 127 + vertical;
            hashCode = hashCode * 127 + indent;
            hashCode = hashCode * 127 + wrapText;
            hashCode = hashCode * 127 + shrinkToFit;
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Alignment alignment = obj as Alignment;
            if (alignment == null) return false;
            if (horizontal == alignment.horizontal
                && vertical == alignment.vertical
                && indent == alignment.indent
                && wrapText == alignment.wrapText
                && shrinkToFit == alignment.shrinkToFit)
                return true;
            return false;
        }

        public string CSVString()
            => $"{horizontal},{vertical},{indent},{wrapText},{shrinkToFit}";
    }

    class Border
    {
        // 0-13 None is 10
        byte left, top, right, bottom;

        public Border(IXLBorder border)
        {
            left = (byte)(border.LeftBorder == XLBorderStyleValues.None ? 0 : 1);
            top = (byte)(border.TopBorder == XLBorderStyleValues.None ? 0 : 1);
            right = (byte)(border.RightBorder == XLBorderStyleValues.None ? 0 : 1);
            bottom = (byte)(border.BottomBorder == XLBorderStyleValues.None ? 0 : 1);
        }

        public override int GetHashCode()
        {
            int hashCode = left;
            hashCode = hashCode * 127 + top;
            hashCode = hashCode * 127 + right;
            hashCode = hashCode * 127 + bottom;
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Border border = obj as Border;
            if (border == null) return false;
            if (left == border.left
                && top == border.top
                && right == border.right
                && bottom == border.bottom)
                return true;
            return false;
        }

        public static string CSVTitle
            = "bdr-left,bdr-top,bdr-right,bdr-bottom";

        public string ToCSV()
            => $"{left},{top},{right},{bottom}";

        public string CSVString()
            => $"{left},{top},{right},{bottom}";
    }

    class Fill
    {
        // 0-18 None is 17, solid is 18
        byte pattern;
        public int foreColor, backColor;
        public int foreIndex, backIndex;

        public Fill(IXLFill fill)
        {
            pattern = (byte)(fill.PatternType == XLFillPatternValues.None ? 0 : 1);
            foreColor = fill.PatternColor.Color.ToArgb();
            backColor = fill.BackgroundColor.Color.ToArgb();
        }

        public override int GetHashCode()
        {
            int hashCode = pattern;
            hashCode = hashCode * 127 + foreColor;
            hashCode = hashCode * 127 + backColor;
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Fill fill = obj as Fill;
            if (fill == null) return false;
            if (pattern == fill.pattern
                && foreColor == fill.foreColor
                && backColor == fill.backColor)
                return true;
            return false;
        }

        public static string CSVTitle
            = "fill-ptn,fill-fcolor,fill-bcolor";

        public string ToCSV()
            => $"{pattern},{foreColor},{backColor}";

        public string CSVString()
            => $"{pattern},{foreIndex},{backIndex}";
    }

    class Font
    {
        byte bold;
        byte italic;
        byte strike;
        // 0-4 None is 2
        byte line;
        // 0-2 Baseline is 0
        byte vertical;
        // 0-5 NotApplicable is 0
        byte familyNumbering;

        public double size;
        public int color;
        public string name;

        public int colorIndex, nameIndex;

        public Font(IXLFont font)
        {
            bold = (byte)(font.Bold ? 1 : 0);
            italic = (byte)(font.Italic ? 1 : 0);
            strike = (byte)(font.Strikethrough ? 1 : 0);
            line = (byte)(font.Underline == XLFontUnderlineValues.None ? 0 : 1);
            vertical = (byte)(font.VerticalAlignment == XLFontVerticalTextAlignmentValues.Baseline ? 0 : 1);
            familyNumbering = (byte)(font.FontFamilyNumbering == XLFontFamilyNumberingValues.NotApplicable ? 0 : 1);
            color = font.FontColor.Color.ToArgb();
            size = font.FontSize;
            name = font.FontName;
        }

        public override int GetHashCode()
        {
            int hashCode = bold;
            hashCode = hashCode * 127 + italic;
            hashCode = hashCode * 127 + strike;
            hashCode = hashCode * 127 + line;
            hashCode = hashCode * 127 + vertical;
            hashCode = hashCode * 127 + familyNumbering;
            hashCode = hashCode * 127 + color;
            hashCode = hashCode * 127 + size.GetHashCode();
            hashCode = hashCode * 127 + name.GetHashCode();
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Font font = obj as Font;
            if (font == null) return false;
            if (bold == font.bold
                && italic == font.italic
                && strike == font.strike
                && line == font.line
                && vertical == font.vertical
                && familyNumbering == font.familyNumbering
                && color == font.color
                && size == font.size
                && name == font.name)
                return true;
            return false;
        }

        public static string CSVTitle
            = "fnt-bold,fnt-italic,fnt-strike,fnt-line,fnt-ver,fnt-color,fnt-size,fnt-name";

        public string ToCSV()
            => $"{bold},{italic},{strike},{line},{vertical},{color},{size},{nameIndex}";

        public string CSVString()
            => $"{bold},{italic},{strike},{line},{vertical},{colorIndex},{size},{nameIndex}";
    }
}
