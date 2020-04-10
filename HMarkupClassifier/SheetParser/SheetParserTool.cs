using ClosedXML.Excel;
using HMarkupClassifier.MarkParser;
using System.IO;


namespace HMarkupClassifier.SheetParser
{
    static class SheetParserTool
    {
        public static XSheet ParseSheet(IXLWorksheet worksheet)
        {
            if (worksheet == null || worksheet.IsEmpty())
                throw new ParseFailException("Sheet is null or Empty.");
            return new XSheet(worksheet);
        }

        public static void WriteCSV(string path, XSheet xSheet)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.WriteLine($"type,row,col,{XCell.CSVTitle}");
                int type = 0;
                for (int row = xSheet.top; row <= xSheet.bottom; row++)
                {
                    for (int col = xSheet.left; col <= xSheet.right; col++)
                    {
                        writer.WriteLine($"{type},{row},{col},{xSheet.cells[(row, col)]}");
                    }
                }
            }
        }

        public static void WriteCSV(string path, XSheet xSheet, YSheet ySheet)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.WriteLine($"type,row,col,{XCell.CSVTitle}");
                int type = 0;
                for (int row = xSheet.top; row <= xSheet.bottom; row++)
                {
                    for (int col = xSheet.left; col <= xSheet.right; col++)
                    {
                        type = ySheet.GetCellType(row, col);
                        writer.WriteLine($"{type},{row},{col},{xSheet.cells[(row,col)]}");
                    }
                }
            }
        }
    }
}
