using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class Style
    {
        int count = 0;
        bool hasPrefix;
        int numberFormat;
        public Alignment alignment;
        public Border border;
        public Fill fill;
        public Font font;

        public Style(IXLStyle style)
        {
            hasPrefix = style.IncludeQuotePrefix;
            numberFormat = style.NumberFormat.NumberFormatId;
            alignment = new Alignment(style.Alignment);
            border = new Border(style.Border);
            fill = new Fill(style.Fill);
            font = new Font(style.Font);
        }

        public override int GetHashCode()
        {
            int hashCode = numberFormat;
            hashCode = hashCode * 127 + hasPrefix.GetHashCode();
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
    }

    class Alignment
    {
        // 0-7 General is 4
        int horizontal;
        // 0-4 Bottom is 0
        int vertical;
        int indent;
        bool wrapText;
        bool shrinkToFit;

        public Alignment(IXLAlignment alignment)
        {
            horizontal = (int)alignment.Horizontal;
            vertical = (int)alignment.Vertical;
            indent = alignment.Indent;
            wrapText = alignment.WrapText;
            shrinkToFit = alignment.ShrinkToFit;
        }

        public override int GetHashCode()
        {
            int hashCode = horizontal;
            hashCode = hashCode * 127 + vertical;
            hashCode = hashCode * 127 + indent;
            hashCode = hashCode * 127 + wrapText.GetHashCode();
            hashCode = hashCode * 127 + shrinkToFit.GetHashCode();
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
    }

    class Border
    {
        // 0-13 None is 0
        int left, top, right, bottom;

        public Border(IXLBorder border)
        {
            left = (int)border.LeftBorder;
            top = (int)border.TopBorder;
            right = (int)border.RightBorder;
            bottom = (int)border.BottomBorder;
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
    }

    class Fill
    {
        // 0-18 None is 17, solid is 18
        int pattern;
        int foreColor, backColor;

        public Fill(IXLFill fill)
        {
            pattern = (int)fill.PatternType;
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
    }

    class Font
    {
        bool isBold;
        bool isItalic;
        bool strikethrough;
        // 0-4 None is 2
        int underline;
        // 0-2 Baseline is 0
        int vertical;
        // 0-5 NotApplicable is 0
        int familyNumbering;

        int color;
        double size;
        string name;

        public Font(IXLFont font)
        {
            isBold = font.Bold;
            isItalic = font.Italic;
            strikethrough = font.Strikethrough;
            underline = (int)font.Underline;
            vertical = (int)font.VerticalAlignment;
            familyNumbering = (int)font.FontFamilyNumbering;
            color = font.FontColor.Color.ToArgb();
            size = font.FontSize;
            name = font.FontName;
        }

        public override int GetHashCode()
        {
            int hashCode = isBold.GetHashCode();
            hashCode = hashCode * 127 + isItalic.GetHashCode();
            hashCode = hashCode * 127 + strikethrough.GetHashCode();
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
            if (isBold == font.isBold
                && isItalic == font.isItalic
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
    }
}
