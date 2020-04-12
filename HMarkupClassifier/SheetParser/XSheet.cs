using ClosedXML.Excel;
using HMarkupClassifier.MarkParser;
using HMarkupClassifier.SheetParser.XPosition;
using System;
using System.Collections.Generic;
using System.IO;

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
        public XStyle sheetStyle;
        public Dictionary<(int, int), XCell> cells = new Dictionary<(int, int), XCell>();

        public XSheet(IXLWorksheet worksheet)
        {
            sheetStyle = xFormatFactory.GetXStyle(worksheet.Style);

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

        public XSide GetXSide(int row, int col, int rstep, int cstep)
        {
            int i = 1;
            while (true)
            {
                if (!cells.ContainsKey((row + i * rstep, col + i * cstep)))
                    return new XSide(cells[(row, col)], sheetStyle, i);
                else if (cells[(row + i * rstep, col + i * cstep)].HasValue == 1)
                    return new XSide(cells[(row, col)], cells[(row + i * rstep, col + i * cstep)], i);
                i++;
            }
        }

        public void WriteCSV(string path, YSheet ySheet)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.Write($"type,row,col,{XCell.CSVTitle}");
                for (int i = 0; i < 4; i++)
                    writer.Write($",{XSide.CSVTitle(i)}");
                writer.WriteLine();
                int type;
                for (int row = top; row <= bottom; row++)
                {
                    for (int col = left; col <= right; col++)
                    {
                        type = ySheet.GetCellType(row, col);
                        writer.Write($"{type},{row},{col},{cells[(row, col)]}");
                        writer.Write($",{GetXSide(row, col, 0, -1)}");
                        writer.Write($",{GetXSide(row, col, -1, 0)}");
                        writer.Write($",{GetXSide(row, col, 0, 1)}");
                        writer.Write($",{GetXSide(row, col, 1, 0)}");
                        writer.WriteLine();
                    }
                }
            }
        }

    }
}
