﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using ClosedXML.Excel;
using HMarkupClassifier.MarkParser;
using HMarkupClassifier.RangeParser;
using HMarkupClassifier.SheetParser.Styles;

namespace HMarkupClassifier.SheetParser
{
    class SheetFeature
    {
        public SheetInfo info;
        private XCellAbd[,] cells;

        private SheetFeature(IXLWorksheet sheet)
        {
            info = new SheetInfo(sheet);
            IXLRange usedCells = sheet.RangeUsed();
            cells = new XCellAbd[info.NumCol, info.NumRow];
            if (info.NumCol * info.NumRow > 1000)
            {
                throw new ParseFailException("Too large range.");
            }
            foreach (var cell in usedCells.Cells())
            {
                int col = cell.Address.ColumnNumber - info.left;
                int row = cell.Address.RowNumber - info.top;
                cells[col, row] = new XCellAbd(cell, info);
            }
            SetFormulaReference();
            info.SetIndex();
            for (int i = 0; i < info.NumCol; i++)
                for (int j = 0; j < info.NumRow; j++)
                {
                    if (i != 0) cells[i, j].SetNeighbors(0, cells[i - 1, j]);
                    if (j != 0) cells[i, j].SetNeighbors(1, cells[i, j - 1]);
                    if (i != info.NumCol - 1) cells[i, j].SetNeighbors(2, cells[i + 1, j]);
                    if (j != info.NumRow - 1) cells[i, j].SetNeighbors(3, cells[i, j + 1]);
                    for (int k = 0; k < info.NumCol; k++)
                        if (i - 1 - k >= 0 && cells[i - 1 - k, j].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(0, cells[i - 1 - k, j]);
                            break;
                        }
                    for (int k = 0; k < info.NumRow; k++)
                        if (j - 1 - k >= 0 && cells[i, j - 1 - k].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(1, cells[i, j - 1 - k]);
                            break;
                        }
                    for (int k = 0; k < info.NumCol; k++)
                        if (i + 1 + k <= info.NumCol - 1 && cells[i + 1 + k, j].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(2, cells[i + 1 + k, j]);
                            break;
                        }
                    for (int k = 0; k < info.NumRow; k++)
                        if (j + 1 + k <= info.NumRow - 1 && cells[i, j + 1 + k].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(3, cells[i, j + 1 + k]);
                            break;
                        }
                }
        }

        public static SheetFeature ParseSheet(IXLWorksheet sheet)
        {
            if (sheet == null || sheet.IsEmpty())
                throw new ParseFailException("Sheet is null or Empty.");
            return new SheetFeature(sheet);
        }

        public void WriteIntoCSV(string path, YSheet markSheet = null)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.WriteLine($"type,{XCellAbd.csvTitle}");
                if (markSheet != null)
                    foreach (var cell in cells)
                    {
                        int type = markSheet.GetCellType(cell.col, cell.row);
                        writer.WriteLine($"{type},{cell.CSVString()}");
                    }
                else
                    foreach (var cell in cells)
                        writer.WriteLine($"0,{cell.CSVString()}");
            }
        }

        public void WriteIntoCSV(string path, SheetMark mark = null)
        {
            using (StreamWriter writer = new StreamWriter(path))
            //using (StreamWriter map = new StreamWriter(path.Substring(0, path.LastIndexOf(".csv")) + ".cellmap"))
            {
                writer.WriteLine($"type,{XCellAbd.csvTitle}");
                if (mark != null)
                    foreach (var cell in cells)
                    {
                        int type = mark.GetCellType(cell.col, cell.row);
                        writer.WriteLine($"{type},{cell.CSVString()}");
                    }
                else
                    foreach (var cell in cells)
                        writer.WriteLine($"{3},{cell.CSVString()}");
            }
        }

        public void SetFormulaReference()
        {
            foreach (var temp in info.formulas)
            {
                var formula = temp.ToUpper();
                foreach (Match match in Formula.A1Regex.Matches(formula))
                {
                    int col = Utils.ParseColumn(match.Groups["Col"].Value);
                    int row = Convert.ToInt32(match.Groups["Row"].Value);
                    if(info.ValidAddress(col, row))
                        cells[col - info.left, row - info.top].isReferenced = 1;
                }
                foreach (Match match in Formula.A1A1Regex.Matches(formula))
                {
                    int left = Utils.ParseColumn(match.Groups["Left"].Value);
                    int top = Convert.ToInt32(match.Groups["Top"].Value);
                    int right = Utils.ParseColumn(match.Groups["Right"].Value);
                    int bottom = Convert.ToInt32(match.Groups["Bottom"].Value);
                    for (int i = left; i <= right; i++)
                        for (int j = top; j <= bottom; j++)
                            if(info.ValidAddress(i, j))
                                cells[i - info.left, j - info.top].isReferenced = 1;
                }
            }
        }

        public override string ToString() => info.ToString();
    }

    class SheetInfo
    {
        public int left, top, right, bottom;
        public int NumCol => right - left + 1;
        public int NumRow => bottom - top + 1;

        public double width, height;

        public List<string> formulas = new List<string>();

        public List<Style> styles = new List<Style>();
        public List<Alignment> alignments = new List<Alignment>();
        public List<Border> borders = new List<Border>();
        public List<Fill> fills = new List<Fill>();
        public List<Font> fonts = new List<Font>();

        public bool ValidAddress(int col, int row)
            => !(col < left || col > right || row < top || row > bottom);


        public SheetInfo(IXLWorksheet sheet)
        {
            Style style = new Style(sheet.Style);
            styles.Add(new Style(sheet.Style));
            alignments.Add(style.Alignment);
            borders.Add(style.Border);
            fills.Add(style.Fill);
            fonts.Add(style.Font);

            IXLRange usedCells = sheet.RangeUsed();
            left = usedCells.RangeAddress.FirstAddress.ColumnNumber;
            top = usedCells.RangeAddress.FirstAddress.RowNumber;
            right = usedCells.RangeAddress.LastAddress.ColumnNumber;
            bottom = usedCells.RangeAddress.LastAddress.RowNumber;
            width = sheet.ColumnWidth;
            height = sheet.RowHeight;
        }

        public Style GetStyle(IXLStyle xlStyle)
        {
            Style style = new Style(xlStyle);
            int index = styles.IndexOf(style);
            if (index >= 0) style = styles[index];
            else
            {
                styles.Add(style);
                if (!alignments.Contains(style.Alignment)) alignments.Add(style.Alignment);
                if (!borders.Contains(style.Border)) borders.Add(style.Border);
                if (!fills.Contains(style.Fill)) fills.Add(style.Fill);
                if (!fonts.Contains(style.Font)) fonts.Add(style.Font);
            }
            style.count++;
            return style;
        }

        public void SetIndex()
        {
            Dictionary<string, int> fontNames = new Dictionary<string, int>();
            Dictionary<int, int> colors = new Dictionary<int, int>();
            foreach (var style in styles)
            {
                if (fontNames.ContainsKey(style.Font.Name)) fontNames[style.Font.Name] += style.count;
                else fontNames.Add(style.Font.Name, style.count);
                if (colors.ContainsKey(style.Font.Color)) colors[style.Font.Color] += style.count;
                else colors.Add(style.Font.Color, style.count);
                if (colors.ContainsKey(style.Fill.PtnColor)) colors[style.Fill.PtnColor] += style.count;
                else colors.Add(style.Fill.PtnColor, style.count);
                if (colors.ContainsKey(style.Fill.BckColor)) colors[style.Fill.BckColor] += style.count;
                else colors.Add(style.Fill.BckColor, style.count);
            }
            var fontSorted = (from entry in fontNames orderby entry.Value descending select entry.Key).ToList();
            var colorSorted = (from entry in colors orderby entry.Value descending select entry.Key).ToList();
            foreach (var style in styles)
            {
                style.Font.NameIndex = fontSorted.IndexOf(style.Font.Name);
                style.Font.ColorIndex = colorSorted.IndexOf(style.Font.Color);
                style.Fill.PtnIndex = colorSorted.IndexOf(style.Fill.PtnColor);
                style.Fill.BckIndex = colorSorted.IndexOf(style.Fill.BckColor);
            }
        }

        public override string ToString()
            => $"[R{top}C{left}:R{bottom}C{right}] Styles{styles.Count}";
    }
}
