using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using ClosedXML.Excel;

namespace HMarkupClassifier
{
    public static class ForTest
    {
        public static string bookfile = "D:\\Temp\\fortest.xlsx";

        public static void Test()
        {
            if (File.Exists(bookfile))
            {
                using (XLWorkbook workbook = new XLWorkbook(bookfile))
                {
                    var sheet = workbook.Worksheets.Worksheet(1);
                    foreach (var cell in sheet.CellsUsed())
                    {
                        Console.WriteLine(cell.Address.ToString() + "\t" + cell.Style.NumberFormat.NumberFormatId.ToString());
                        if (cell.Address.ColumnNumber == 3)
                        {
                            cell.Style.NumberFormat.NumberFormatId = 2;
                        }
                    }
                    workbook.Save();
                }
            }
            else
                Console.WriteLine("Not Exsit");
        }
    }
}
