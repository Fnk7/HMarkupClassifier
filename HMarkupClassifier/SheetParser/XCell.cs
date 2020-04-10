using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HMarkupClassifier.SheetParser
{
    class XCell
    {
        public int HasValue;
        public int Merged;
        public XFormat Format;
        public XContent Content;
        public XFormula Formula;

        public XCell(IXLCell cell)
        {
            HasValue = cell.IsEmpty() ? 0 : 1;
            Merged = cell.IsMerged() ? 1 : 0;
        }

        public static string CSVTitle
            = $"hasvalue,merged,{XFormat.CSVTitle},{XContent.CSVTitle},{XFormula.CSVTitle}";

        public override string ToString()
        {
            return base.ToString();
        }
    }
}
