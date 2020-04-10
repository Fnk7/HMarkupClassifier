using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class XFormat
    {
        // Cell Style
        public XStyle Style;
        public int DataType;
        public int HasHyperlink;
        public int HasComment;

        public XFormat(IXLCell cell)
        {
            DataType = (int)cell.DataType;
            HasHyperlink = cell.HasHyperlink ? 1 : 0;
            HasComment = cell.HasComment ? 1 : 0;
        }

        public static string CSVTitle
            = $"datatype,{XStyle.CSVTitle},fmt-link,fmt-comment";

        public override string ToString()
        {
            return $"{DataType},{Style},{HasHyperlink},{HasComment}";
        }
    }
}
