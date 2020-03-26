using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;
using HMarkupClassifier.RangeParser;

namespace HMarkupClassifier
{
    public static class Utils
    {
        public static void ParseDataset(string dataset, string target)
        {
            if (!Directory.Exists(target))
                Directory.CreateDirectory(target);
            else if (Directory.GetFiles(target).Length != 0)
                return;
            
            var rangeFiles = Directory.GetFiles(dataset, "*.range");
            int sheetIndex = 1;
            using (StreamWriter infoWriter = new StreamWriter(Path.Combine(target, "Info.txt")))
            {
                foreach (var rangeFile in rangeFiles)
                {
                    string workbookFile = rangeFile.Substring(0, rangeFile.LastIndexOf(".range")) + ".xlsx";
                    if (File.Exists(workbookFile))
                    {
                        infoWriter.WriteLine(rangeFile);
                        infoWriter.WriteLine(workbookFile);
                        Dictionary<string, SheetMark> sheetMarks = SheetMark.ParseRange(rangeFile);
                        using (XLWorkbook book = new XLWorkbook(workbookFile))
                        {
                            foreach (var sheet in book.Worksheets)
                                if (sheetMarks.ContainsKey(sheet.Name))
                                    try
                                    {
                                        SheetFeature feature = SheetFeature.ParseSheet(sheet);
                                        string path = Path.Combine(target, $"Sheet{sheetIndex:D3}.csv");
                                        feature.WriteIntoCSV(path, sheetMarks[sheet.Name]);
                                        infoWriter.WriteLine($"{sheet.Name}\tSheet{sheetIndex:D3}.csv");
                                        sheetIndex++;
                                    }
                                    catch (Exception ex)
                                    {
                                        infoWriter.Write(sheet.Name);
                                        infoWriter.Write("\t");
                                        infoWriter.WriteLine(ex.Message);
                                    }
                        }
                        infoWriter.WriteLine();
                    }
                }
            }
        }



        public static void Test()
        {
            Console.WriteLine("Test start...");
            using (XLWorkbook book = new XLWorkbook("D:\\Workspace\\HMarkupDataset\\Annotated\\02rise.xlsx"))
            {
                var sheet = book.Worksheet("Companies");
                foreach (var cell in sheet.RangeUsed(XLCellsUsedOptions.All).Cells())
                {
                    Console.WriteLine($"{cell.Address.ToString()}\t{cell.Style.NumberFormat.Format}");
                }
            }
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
                                List<Style> styles = feature.info.styles;
                                List<Alignment> alignments = feature.info.alignments;
                                List<Border> borders = feature.info.borders;
                                List<Fill> fills = feature.info.fills;
                                List<Font> fonts = feature.info.fonts;
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
            if (!Directory.Exists("D:\\Statistic2"))
                Directory.CreateDirectory("D:\\Statistic2");
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
                workbook.SaveAs(Path.Combine("D:\\Statistic2", name + ".xlsx"));
            }
        }
    }
}
