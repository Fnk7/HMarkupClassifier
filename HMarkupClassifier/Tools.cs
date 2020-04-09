using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

using HMarkupClassifier.MarkParser;

namespace HMarkupClassifier
{
    public static class Tools
    {
        private static readonly string xlsx = ".xlsx";
        private static readonly string csv = ".csv";


        public static void ParseMarkDst(string markDst, string CSVDst)
        {
            if (!Directory.Exists(CSVDst))
                Directory.CreateDirectory(CSVDst);
            else if (Directory.GetFiles(CSVDst).Length != 0)
                throw new Exception("CSV Dataset Not Empty.");

            int sheetIndex = 1;
            var markFiles = Directory.GetFiles(markDst, "*.mark");
            var rangeFiles = Directory.GetFiles(markDst, "*.range");
            var files = new List<string>(markFiles).Concat(rangeFiles);
            using (StreamWriter infoWriter = new StreamWriter(Path.Combine(CSVDst, "Info.txt")))
            {
                foreach (var markFile in files)
                {
                    infoWriter.WriteLine(markFile);
                    var workbookFile = markFile.Substring(0, markFile.LastIndexOf('.')) + xlsx;
                    if (File.Exists(workbookFile))
                    {
                        infoWriter.WriteLine(workbookFile);
                        Dictionary<string, YSheet> markSheets = MarkParserTool.Parse(markFile);
                        using (XLWorkbook book = new XLWorkbook(workbookFile))
                            foreach (var sheet in book.Worksheets)
                                if (markSheets.ContainsKey(sheet.Name))
                                    try
                                    {
                                        SheetFeature feature = SheetFeature.ParseSheet(sheet);
                                        string CSVPath = Path.Combine(CSVDst, $"Sheet{sheetIndex:D3}{csv}");
                                        feature.WriteIntoCSV(CSVPath, markSheets[sheet.Name]);
                                        infoWriter.WriteLine($"{sheet.Name}\tSheet{sheetIndex:D3}{csv}");
                                        sheetIndex++;
                                    }
                                    catch (ParseFailException ex)
                                    {
                                        infoWriter.WriteLine($"{sheet.Name}\t{ex.Message}");
                                        Console.WriteLine($"{sheet.Name}\t{ex.Message}");
                                    }
                        infoWriter.WriteLine();
                    }
                }
            }
        }



    }
}
