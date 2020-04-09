using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class XCell
    {
        public int col, row;
        public Format format;
        public Content content;

        public XCell(IXLCell cell, Statistic statistic)
        {
            col = cell.Address.ColumnNumber;
            row = cell.Address.RowNumber;

            format = new Format(cell, statistic);
            content = new Content(cell, statistic);

            
        }
    }
}
