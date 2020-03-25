using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;

namespace HMarkupClassifier
{
    public class Utils
    {
        public static void Test()
        {
            Console.WriteLine("Test start...");
            XLWorkbook book = new XLWorkbook("D:\\02rise.xlsx");
            foreach (var sheet in book.Worksheets)
            {
                try
                {
                    SheetFeature feature = SheetFeature.ParseSheet(sheet);
                    Console.WriteLine(feature.ToString());
                }catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.StackTrace);
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
