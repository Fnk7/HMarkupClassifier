using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class Style
    {
        public int count = 0;

        public byte hasPrefix;
        public byte numberFormat;
        public Alignment alignment;
        public Border border;
        public Fill fill;
        public Font font;

        private string csv = null;

        public Style(IXLStyle style)
        {
            hasPrefix = (byte)(style.IncludeQuotePrefix ? 1 : 0);
            numberFormat = (byte)(style.NumberFormat.NumberFormatId == 0 ? 0 : 1);
            alignment = new Alignment(style.Alignment);
            border = new Border(style.Border);
            fill = new Fill(style.Fill);
            font = new Font(style.Font);
        }

        public override int GetHashCode()
        {
            int hashCode = hasPrefix;
            hashCode = hashCode * 127 + numberFormat;
            hashCode = hashCode * 127 + alignment.GetHashCode();
            hashCode = hashCode * 127 + border.GetHashCode();
            hashCode = hashCode * 127 + fill.GetHashCode();
            hashCode = hashCode * 127 + font.GetHashCode();
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Style style = obj as Style;
            if (obj == null) return false;
            if (hasPrefix == style.hasPrefix
                && numberFormat == style.numberFormat
                && alignment.Equals(style.alignment)
                && border.Equals(style.border)
                && fill.Equals(style.fill)
                && font.Equals(style.font))
                return true;
            return false;
        }

        public static string csvTitle
            = $"Prefix,Format,{Alignment.csvTitle},{Border.csvTitle},{Fill.csvTitle},{Font.csvTitle}";

        public string CSVString()
        {
            if (csv == null)
                csv = $"{hasPrefix},{numberFormat},{alignment.CSVString()}," +
                    $"{border.CSVString()},{fill.CSVString()},{font.CSVString()}";
            return csv;
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

        public static string csvTitle
            = "Horizontal,Vertical-ALI,Indent,WrapText,ShrinkToFik";

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

        public static string csvTitle
            = "BorderLeft,BorderTop,BorderRight,BorderBottom";

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

        public static string csvTitle
            = "Fill,ForeColor,BackColor";

        public string CSVString()
            => $"{pattern},{foreIndex},{backIndex}";
    }

    class Font
    {
        byte bold;
        byte italic;
        byte strikethrough;
        // 0-4 None is 2
        byte underline;
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
            strikethrough = (byte)(font.Strikethrough ? 1 : 0);
            underline = (byte)(font.Underline == XLFontUnderlineValues.None ? 0 : 1);
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
            hashCode = hashCode * 127 + strikethrough;
            hashCode = hashCode * 127 + underline;
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
                && strikethrough == font.strikethrough
                && underline == font.underline
                && vertical == font.vertical
                && familyNumbering == font.familyNumbering
                && color == font.color
                && size == font.size
                && name == font.name)
                return true;
            return false;
        }

        public static string csvTitle
            = "Bold,Italic,Strikethrough,Underline,Vertical-Font,FamilyNumbering,FontColor,FontSize,FontName";

        public string CSVString()
            => $"{bold},{italic},{strikethrough},{underline},{vertical},{familyNumbering},{colorIndex},{size},{nameIndex}";
    }
}
