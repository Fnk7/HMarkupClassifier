using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class SheetFeature
    {
        private int firstCol, firstRow, lastCol, lastRow;
        private CellFeature[,] cells;
        public List<Style> styles = new List<Style>();

        private SheetFeature(IXLWorksheet sheet)
        {
            styles.Add(new Style(sheet.Style));
            IXLRange usedCells = sheet.RangeUsed(XLCellsUsedOptions.All);
            firstCol = usedCells.RangeAddress.FirstAddress.ColumnNumber;
            firstRow = usedCells.RangeAddress.FirstAddress.RowNumber;
            lastCol = usedCells.RangeAddress.LastAddress.ColumnNumber;
            lastRow = usedCells.RangeAddress.LastAddress.RowNumber;
            cells = new CellFeature[lastCol - firstCol + 1, lastRow - firstRow + 1];
            foreach (var cell in usedCells.Cells())
            {
                int col = cell.Address.ColumnNumber - firstCol;
                int row = cell.Address.RowNumber - firstRow;
                cells[col, row] = new CellFeature(cell, styles);
            }
        }

        public static SheetFeature ParseSheet(IXLWorksheet sheet)
        {
            if (sheet == null || sheet.IsEmpty())
                throw new ParseFailException("Sheet is null or Empty.");
            return new SheetFeature(sheet);
        }

        public override string ToString()
        {
            return string.Format("[R{0}C{1}:R{2}C{3}] styles:{4}", 
                firstRow, firstCol, lastRow, lastCol, styles.Count);
        }
    }
}
