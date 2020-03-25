using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;
using HMarkupClassifier.RangeParser;

namespace HMarkupClassifier
{
    public static class Utils
    {
        public static void ParseAnnotated(string dataset, string target)
        {
            dataset = Path.GetFullPath(dataset);
            target = Path.GetFullPath(target);
            if (!Directory.Exists(target))
                Directory.CreateDirectory(target);
            else if (Directory.GetFiles(target).Length != 0)
                return;
            var rangeFiles = Directory.GetFiles(dataset, "*.range");
            foreach (var rangeFile in rangeFiles)
            {
                string workbookFile = rangeFile.Substring(0, rangeFile.LastIndexOf(".range")) + ".xlsx";
                if (File.Exists(workbookFile))
                {
                    Dictionary<string, SheetMark> sheetMarks = SheetMark.ParseRange(rangeFile);
                    using (XLWorkbook book = new XLWorkbook(workbookFile))
                    {
                        foreach (var sheetMark in sheetMarks)
                        {
                            try
                            {
                                var sheet = book.Worksheet(sheetMark.Key);
                                SheetFeature sheetFeature = SheetFeature.ParseSheet(sheet);

                            }
                            catch (Exception ex)
                            {
                                Console.Write($"In {workbookFile} :");
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
            }
        }


        private static void WriteIntoCSV(SheetFeature sheetFeature, SheetMark sheetMark = null, string target, string sheetName)
        {
            sheetName = GetValidFileName(sheetName);
            if(sheetName == "" || File.Exists(Path.Combine(target, sheetName)))
            {
                int temp = 1;
                string tempName;
                do
                {
                    tempName = sheetName + $"({temp})";
                    temp++;
                } while (File.Exists(Path.Combine(target, tempName)));
                sheetName = tempName;
                
            }

        }

        private static string GetValidFileName(string fileName)
        {
            StringBuilder builder = new StringBuilder(fileName);
            foreach (var ch in Path.GetInvalidFileNameChars())
                builder.Replace(ch.ToString(), string.Empty);
            return builder.ToString();
        }

        public static void Test()
        {
            Console.WriteLine("Test start...");
            XLWorkbook workbook = new XLWorkbook("D:\\test.xlsx");

            Console.WriteLine("Test end...");
        }

        public static void ShowFreq()
        {
            var rangeFiles = Directory.GetFiles("D:\\Annotated", "*.range");
            int sheetCnt = 0;
            Dictionary<int, int> styleCnt = new Dictionary<int, int>();
            Dictionary<int, int> borderCnt = new Dictionary<int, int>();
            Dictionary<int, int> fillCnt = new Dictionary<int, int>();
            Dictionary<int, int> fontCnt = new Dictionary<int, int>();
            Dictionary<int, int> alignmentCnt = new Dictionary<int, int>();
            Console.WriteLine("Start....");
            foreach (var rangeFile in rangeFiles)
            {
                string workbookFile = rangeFile.Substring(0, rangeFile.LastIndexOf(".range")) + ".xlsx";
                if (File.Exists(workbookFile))
                {
                    using (XLWorkbook book = new XLWorkbook(workbookFile))
                    {
                        foreach (var sheet in book.Worksheets)
                        {
                            try
                            {
                                sheetCnt++;
                                SheetFeature feature = SheetFeature.ParseSheet(sheet);
                                Console.WriteLine(feature.ToString());
                                List<Style> styles = feature.styles;
                                List<Border> borders = new List<Border>();
                                List<Fill> fills = new List<Fill>();
                                List<Font> fonts = new List<Font>();
                                List<Alignment> alignments = new List<Alignment>();
                                foreach (var style in styles)
                                {
                                    if (!borders.Contains(style.border)) borders.Add(style.border);
                                    if (!fills.Contains(style.fill)) fills.Add(style.fill);
                                    if (!fonts.Contains(style.font)) fonts.Add(style.font);
                                    if (!alignments.Contains(style.alignment)) alignments.Add(style.alignment);
                                }
                                if (styleCnt.ContainsKey(styles.Count)) styleCnt[styles.Count]++;
                                else styleCnt.Add(styles.Count, 1);
                                if (borderCnt.ContainsKey(borders.Count)) borderCnt[borders.Count]++;
                                else borderCnt.Add(borders.Count, 1);
                                if (fillCnt.ContainsKey(fills.Count)) fillCnt[fills.Count]++;
                                else fillCnt.Add(fills.Count, 1);
                                if (fontCnt.ContainsKey(fonts.Count)) fontCnt[fonts.Count]++;
                                else fontCnt.Add(fonts.Count, 1);
                                if (alignmentCnt.ContainsKey(alignments.Count)) alignmentCnt[alignments.Count]++;
                                else alignmentCnt.Add(alignments.Count, 1);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine("Total Sheets: {0}", sheetCnt);
            if (!Directory.Exists("D:\\Statistic"))
                Directory.CreateDirectory("D:\\Statistic");
            PrintCount(styleCnt, "Style", sheetCnt);
            PrintCount(borderCnt, "Border", sheetCnt);
            PrintCount(fillCnt, "Fill", sheetCnt);
            PrintCount(fontCnt, "Font", sheetCnt);
            PrintCount(alignmentCnt, "Alignment", sheetCnt);
            Console.WriteLine("Total Sheets: {0}", sheetCnt);
            Console.ReadLine();
        }

        public static void PrintCount(Dictionary<int, int> statistics, string name, int total)
        {
            var orderedCnt = from entry in statistics orderby entry.Key ascending select entry;
            int sheetCnt = 0;
            Console.Write(name);
            using (XLWorkbook workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("sheet1");
                IXLRow row = workbook.Worksheet("sheet1").FirstRow();
                IXLCell cell = row.FirstCell();
                cell.Value = "Key";
                cell = cell.CellRight();
                cell.Value = "Time";
                cell = cell.CellRight();
                cell.Value = "Total";
                row = row.RowBelow();
                foreach (var time in orderedCnt)
                {
                    Console.WriteLine("\tKey:{0}, count:{1}", time.Key, time.Value);
                    if (sheetCnt >= total * 5 / 6)
                    {
                        cell = row.FirstCell();
                        cell.Value = time.Key;
                        cell = cell.CellRight();
                        cell.Value = total - sheetCnt;
                        cell = cell.CellRight();
                        cell.Value = total;
                        break;
                    }
                    sheetCnt += time.Value;
                    cell = row.FirstCell();
                    cell.Value = time.Key;
                    cell = cell.CellRight();
                    cell.Value = time.Value;
                    cell = cell.CellRight();
                    cell.Value = sheetCnt;
                    row = row.RowBelow();
                }
                Console.WriteLine("Sheet Count: " + sheetCnt);
                workbook.SaveAs(Path.Combine("D:\\Statistic", name + ".xlsx"));
            }
        }
    }
}
