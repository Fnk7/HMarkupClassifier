using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser.Styles
{
    class XStyle
    {
        public int count = 0;

        // TODO Number Format
        public int HasPrefix;
        public int NumFormat;

        public XAlignment Alignment;
        public XBorder Border;
        public XFill Fill;
        public XFont Font;

        public XStyle(IXLStyle style)
        {
            HasPrefix = style.IncludeQuotePrefix ? 1 : 0;
            NumFormat = style.NumberFormat.NumberFormatId;
            Alignment = new XAlignment(style.Alignment);
            Border = new XBorder(style.Border);
            Fill = new XFill(style.Fill);
            Font = new XFont(style.Font);
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
            XStyle style = obj as XStyle;
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
            = $"style-prefix,style-numformat,{XAlignment.CSVTitle},{XBorder.CSVTitle},{XFill.CSVTitle},{XFont.CSVTitle}";

        private string CSV = null;
        public override string ToString()
        {
            if (CSV == null)
                CSV = $"{HasPrefix},{NumFormat},{Alignment},{Border},{Fill},{Font}";
            return CSV;
        }
    }
}
