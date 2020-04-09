using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class Format
    {
        // Cell Style
        public Style style;
        public int merged = 0;
        public int hide = 0;
        public double width, height;

        public Format(IXLCell cell, Statistic statistic)
        {
            style = statistic.GetStyle(cell.Style);
            if (cell.IsMerged()) merged = 1;
            if (cell.WorksheetColumn().IsHidden || cell.WorksheetRow().IsHidden) hide = 1;
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;
        }

        public static string CSVTitle
            => $"{Style.CSVTitle},fmt-merged,fmt-hide,fmt-width,fmt-height";

        public string ToCSV()
            => $"{style.ToCSV()},{merged},{hide},{width},{height}";
    }
}
