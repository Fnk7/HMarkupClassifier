using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HMarkupClassifier.SheetParser.Styles
{
    class Style
    {
        public int count = 0;

        // TODO Number Format
        public int HasPrefix;
        public int NumFormat;

        public Alignment Alignment;
        public Border Border;
        public Fill Fill;
        public Font Font;

        public Style(IXLStyle style)
        {
            HasPrefix = style.IncludeQuotePrefix ? 1 : 0;
            // TODO
            Alignment = new Alignment(style.Alignment);
            Border = new Border(style.Border);
            Fill = new Fill(style.Fill);
            Font = new Font(style.Font);
        }

        public override int GetHashCode()
        {
            int hashCode = HasPrefix;
            hashCode = (hashCode << 3) ^ NumFormat;
            hashCode = (hashCode << 3) ^ Alignment.GetHashCode();
            hashCode = (hashCode << 3) ^ Border.GetHashCode();
            hashCode = (hashCode << 3) ^ Fill.GetHashCode();
            hashCode = (hashCode << 3) ^ Font.GetHashCode();
            return hashCode;
        }

        public override bool Equals(object obj)
        {
            Style style = obj as Style;
            if (obj == null) return false;
            if (HasPrefix == style.HasPrefix
                && NumFormat == style.NumFormat
                && Alignment.Equals(style.Alignment)
                && Border.Equals(style.Border)
                && Fill.Equals(style.Fill)
                && Font.Equals(style.Font))
                return true;
            return false;
        }

        public static string CSVTitle
            = $"style-prefix,style-numformat,{Alignment.CSVTitle},{Border.CSVTitle},{Fill.CSVTitle},{Font.CSVTitle}";

        private string CSV = null;
        public override string ToString()
        {
            if (CSV == null)
                CSV = $"{HasPrefix},{NumFormat},{Alignment},{Border},{Fill},{Font}";
            return CSV;
        }
    }
}
