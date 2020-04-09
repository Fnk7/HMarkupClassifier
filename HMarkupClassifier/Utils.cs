﻿using System;
using System.Collections.Generic;
using System.IO;

using ClosedXML.Excel;
using HMarkupClassifier.SheetParser;
using HMarkupClassifier.RangeParser;


namespace HMarkupClassifier
{
    public static class Utils
    {
        public static string annotatedDataset = "D:\\Workspace\\HMarkupDataset\\Annotated";
        public static string csvDataset = "D:\\Workspace\\HMarkupDataset\\CSV-Final2";
        public static string pythonFile = "D:\\Workspace\\Python\\HMarkup\\main.py";
        public static string pythonModel = csvDataset + "\\forest.model";

        public static int ParseColumn(string col)
        {
            int temp = 0;
            foreach (var c in col)
                temp = temp * 26 + c - 'A' + 1;
            return temp;
        }

        public static bool ParseDataset(string dataset, string target)
        {
            if (!Directory.Exists(target))
                Directory.CreateDirectory(target);
            else if (Directory.GetFiles(target).Length != 0)
                return false;
            
            var rangeFiles = Directory.GetFiles(dataset, "*.range");
            int sheetIndex = 1;
            using (StreamWriter infoWriter = new StreamWriter(Path.Combine(target, "Info.txt")))
            {
                foreach (var rangeFile in rangeFiles)
                {
                    string workbookFile = rangeFile.Substring(0, rangeFile.LastIndexOf(".range")) + ".xlsx";
                    if (File.Exists(workbookFile))
                    {
                        infoWriter.WriteLine(workbookFile);
                        Dictionary<string, SheetMark> sheetMarks = SheetMark.ParseRange(rangeFile);
                        using (XLWorkbook book = new XLWorkbook(workbookFile))
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
                                    catch (ParseFailException ex)
                                    {
                                        infoWriter.WriteLine($"{sheet.Name}\t{ex.Message}");
                                        Console.WriteLine($"{sheet.Name}\t{ex.Message}");
                                    }
                        infoWriter.WriteLine();
                    }
                }
            }
            return true;
        }


        public static int RunPython(string pythonFile, string argument)
        {
            using (System.Diagnostics.Process process = new System.Diagnostics.Process())
            {
                process.StartInfo.FileName = "python";
                process.StartInfo.Arguments = $"{pythonFile} {argument}";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.Start();
                process.WaitForExit();
                var message = process.StandardOutput.ReadToEnd();
                Console.WriteLine(message);
                return process.ExitCode;
            }
        }


        public static void Test()
        {
            Console.WriteLine("Test start...");
            Console.WriteLine(Directory.GetCurrentDirectory());
            Console.WriteLine(Environment.CurrentDirectory);
            Console.WriteLine(AppDomain.CurrentDomain.BaseDirectory);
            Console.WriteLine("Test end...");
        }
    }
}
