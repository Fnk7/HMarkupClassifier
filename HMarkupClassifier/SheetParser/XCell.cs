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
        public XFormula Formula;
        public XContent Content;

        public XCell(IXLCell cell)
        {
            HasValue = cell.IsEmpty() ? 0 : 1;
            Merged = cell.IsMerged() ? 1 : 0;
        }

        public static string CSVTitle
            = $"hasvalue,merged,{XFormat.CSVTitle},{XFormula.CSVTitle},{XContent.CSVTitle}";

        public override string ToString()
        {
            return $"{HasValue},{Merged},{Format},{Formula},{Content}";
        }
    }
}
