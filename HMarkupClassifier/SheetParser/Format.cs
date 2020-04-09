using ClosedXML.Excel;

using HMarkupClassifier.SheetParser.Styles;

namespace HMarkupClassifier.SheetParser
{
    class Format
    {
        // Cell Style
        public Style Style;
        public int Merged = 0;
        public double width, height;

        public Format(IXLCell cell, Statistic statistic)
        {
            Style = statistic.GetStyle(cell.Style);
            if (cell.IsMerged()) Merged = 1;
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;
        }

        public static string CSVTitle
            => $"{Style.CSVTitle},fmt-merged,fmt-width,fmt-height";

        public string ToCSV()
            => $"{Style},{Merged},{width},{height}";
    }
}
