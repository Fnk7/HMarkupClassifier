using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HMarkupClassifier.Predict
{
    public static class Predict
    {
        public static List<(int, int)> PredictHeader(string pythonFile, string clf, string workbookFile, string sheetName, string tempdir = null)
        {
            XSheet xSheet;
            using (XLWorkbook workbook = new XLWorkbook(workbookFile))
            {
                if (!workbook.Worksheets.Contains(sheetName))
                    throw new Exception("Sheet Not Exits");
                var worksheet = workbook.Worksheet(sheetName);
                xSheet = SheetParserTool.ParseSheet(worksheet);
            }
            if (tempdir == null)
                tempdir = Path.GetTempPath();
            if (tempdir == null)
                throw new Exception("Temp folder not exsit");
            string[] argument = new string[3];
            argument[0] = clf;
            argument[1] = Path.Combine(tempdir, "temp_sheet.csv");
            argument[2] = Path.Combine(tempdir, "temp_result.csv");
            xSheet.WriteCSV(argument[1]);
            var code = Utils.RunPython(pythonFile, argument);
            if (code != 0 || !File.Exists(argument[2]))
                throw new Exception("Run Classifier Failed!");
            return GetHeaders(argument[2]);
        }

        public static List<(int, int)> GetHeaders(string resultFile)
        {
            Dictionary<(int, int), int> result = new Dictionary<(int, int), int>();
            try
            {
                using (StreamReader reader = new StreamReader(resultFile))
                {
                    var title = reader.ReadLine();
                    if (title != "type,row,col")
                        throw new Exception("Title Error.");
                    while (reader.ReadLine() is string line)
                    {
                        var values = line.Split(',');
                        int rst = Convert.ToInt32(values[0]);
                        int row = Convert.ToInt32(values[1]);
                        int col = Convert.ToInt32(values[2]);
                        result[(row, col)] = rst;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Result File Error", ex);
            }
            return (from cell in result where cell.Value == 1 select cell.Key).ToList();
        }
    }
}
