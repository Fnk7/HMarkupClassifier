namespace HMarkupClassifier.SheetParser.XPosition
{
    class XSide
    {
        public static readonly string[] N =
            { "side-left-", "side-top-", "side-right-", "side-bottom-"};

        public int Distance;
        public int Inside;
        public int SDataType;
        public int SNumFormat;
        public int DAlignment;
        public int DFill;
        public int DFont;
        public int SFormula;
        public int SContent;

        public XSide(XCell center, XStyle style, int distance)
        {
            Distance = distance;
            Inside = 0;
            SDataType = 0;
            SNumFormat = 0;
            DAlignment = center.Format.Style.Alignment.Equals(style.Alignment) ? 0 : 1;
            DFill = center.Format.Style.Fill.Equals(style.Fill) ? 0 : 1;
            DFont = center.Format.Style.Font.Equals(style.Font) ? 0 : 1;
            SFormula = 0;
            SContent = 0;
        }

        public XSide(XCell center, XCell other, int distance)
        {
            Distance = distance;
            Inside = 1;
            SDataType = center.Format.DataType == other.Format.DataType ? 1 : 0;
            SNumFormat = center.Format.Style.NumFormat == other.Format.Style.NumFormat ? 1 : 0;
            DAlignment = center.Format.Style.Alignment.Equals(other.Format.Style.Alignment) ? 0 : 1;
            DFill = center.Format.Style.Fill.Equals(other.Format.Style.Fill) ? 0 : 1;
            DFont = center.Format.Style.Font.Equals(other.Format.Style.Font) ? 0 : 1;
            SFormula = center.Formula.Similar(other.Formula);
            SContent = center.Content.Similar(other.Content);
        }

        public static string CSVTitle(int s)
            => $"{N[s]}distance," +
               $"{N[s]}inside," +
               $"{N[s]}sdatatype," +
               $"{N[s]}snumformat," +
               $"{N[s]}dalignment," +
               $"{N[s]}dfill," +
               $"{N[s]}dfont," +
               $"{N[s]}sformula," +
               $"{N[s]}scontent";

        public override string ToString()
        {
            return $"{Distance}," +
                   $"{Inside}," +
                   $"{SDataType}," +
                   $"{SNumFormat}," +
                   $"{DAlignment}," +
                   $"{DFill}," +
                   $"{DFont}," +
                   $"{SFormula}," +
                   $"{SContent}";
        }
    }
}
