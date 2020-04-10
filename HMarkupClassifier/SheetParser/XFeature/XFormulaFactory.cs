using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class XFormulaFactory
    {
        public List<string> FormulaA1s = new List<string>();

        public XFormula GetXFormula(IXLCell cell)
        {
            XFormula xFormula = new XFormula();
            if (cell.HasFormula)
                FormulaA1s.Add(cell.FormulaA1);
            xFormula.HasFormula = cell.HasFormula ? 1 : 0;
            xFormula.HasArrayFormula = cell.HasArrayFormula ? 1 : 0;
            return xFormula;
        }

        public void SetFormulaReferenced(Dictionary<(int, int), XCell> cells)
        {
            foreach (string Formula in FormulaA1s)
            {
                foreach (Match match in XFormula.A1Regex.Matches(Formula.ToUpper()))
                {
                    int left = Utils.ParseColumn(match.Groups["Left"].Value);
                    int top = Convert.ToInt32(match.Groups["Top"].Value);
                    if (match.Groups["Right"].Value != string.Empty)
                    {
                        int right = Utils.ParseColumn(match.Groups["Right"].Value);
                        int bottom = Convert.ToInt32(match.Groups["Bottom"].Value);
                        for (int row = top; row <= bottom; row++)
                        {
                            for (int col = left; col <= right; col++)
                            {
                                if (cells.ContainsKey((row, col)))
                                {
                                    cells[(row, col)].Formula.IsReferenced += 1;
                                }
                            }
                        }
                    }
                    else if (cells.ContainsKey((top, left)))
                    {
                        cells[(top, left)].Formula.IsReferenced += 1;
                    }
                }
            }
        }
    }
}
