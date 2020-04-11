using ClosedXML.Excel;
using System.Collections.Generic;

namespace HMarkupClassifier.SheetParser
{
    class XSheet
    {
        public int left, top, right, bottom;
        public int NumCol => right - left + 1;
        public int NumRow => bottom - top + 1;

        public XFormatFactory xFormatFactory = new XFormatFactory();
        public XFormulaFactory xFormulaFactory = new XFormulaFactory();
        public XContentFactory xContentFactory = new XContentFactory();
        public Dictionary<(int, int), XCell> cells = new Dictionary<(int, int), XCell>();

        public XSheet(IXLWorksheet worksheet)
        {
            IXLRange usedCells = worksheet.RangeUsed();
            left = usedCells.RangeAddress.FirstAddress.ColumnNumber;
            top = usedCells.RangeAddress.FirstAddress.RowNumber;
            right = usedCells.RangeAddress.LastAddress.ColumnNumber;
            bottom = usedCells.RangeAddress.LastAddress.RowNumber;

            for (int row = top; row <= bottom; row++)
            {
                for (int col = left; col <= right; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    if (cell.IsMerged())
                    {
                        var fCell = cell.MergedRange().FirstCell();
                        int mrow = fCell.Address.RowNumber;
                        int mcol = fCell.Address.ColumnNumber;
                        if (mrow != row || mcol != col)
                        {
                            cells[(row, col)] = cells[(mrow, mcol)];
                            continue;
                        }
                    }
                    var xCell = new XCell(cell);
                    cells[(row, col)] = xCell;
                    xCell.HasValue = cell.IsEmpty() ? 0 : 1;
                    xCell.Merged = cell.IsMerged() ? 1 : 0;
                    xCell.Format = xFormatFactory.GetXFormat(cell);
                    xCell.Content = xContentFactory.GetXContent(cell);
                    xCell.Formula = xFormulaFactory.GetXFormula(cell);
                }
            }
            xFormatFactory.SetOrderedIndex();
            xFormulaFactory.SetFormulaReferenced(cells);
        }
    }
}
