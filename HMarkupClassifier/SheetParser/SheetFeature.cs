using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using ClosedXML.Excel;

using HMarkupClassifier.RangeParser;

namespace HMarkupClassifier.SheetParser
{
    class SheetFeature
    {
        public SheetInfo info;
        private CellFeature[,] cells;

        private SheetFeature(IXLWorksheet sheet)
        {
            info = new SheetInfo(sheet);
            IXLRange usedCells = sheet.RangeUsed();
            cells = new CellFeature[info.NumCol, info.NumRow];
            foreach (var cell in usedCells.Cells())
            {
                int col = cell.Address.ColumnNumber - info.left;
                int row = cell.Address.RowNumber - info.top;
                cells[col, row] = new CellFeature(cell, info);
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
                            cells[i, j].SetNonEmptyNeighbors(1, cells[i - 1 - k, j]);
                            cells[i, j].leftDistance = k;
                            break;
                        }
                    for (int k = 0; k < info.NumRow; k++)
                        if (j - 1 - k >= 0 && cells[i, j - 1 - k].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(1, cells[i, j - 1 - k]);
                            cells[i, j].topDistance = k;
                            break;
                        }
                    for (int k = 0; k < info.NumCol; k++)
                        if (i + 1 + k <= info.NumCol - 1 && cells[i + 1 + k, j].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(1, cells[i + 1 + k, j]);
                            cells[i, j].rightDistance = k;
                            break;
                        }
                    for (int k = 0; k < info.NumRow; k++)
                        if (j + 1 + k <= info.NumRow - 1 && cells[i, j + 1 + k].empty == 0)
                        {
                            cells[i, j].SetNonEmptyNeighbors(1, cells[i, j + 1 + k]);
                            cells[i, j].bottomDistance = k;
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

        public void WriteIntoCSV(string path, SheetMark mark = null)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.WriteLine($"type,{CellFeature.csvTitle}");
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
        {
            if (col < left || col > right || row < top || row > bottom) return false;
            return true;
        }


        public SheetInfo(IXLWorksheet sheet)
        {
            Style style = new Style(sheet.Style);
            styles.Add(new Style(sheet.Style));
            alignments.Add(style.alignment);
            borders.Add(style.border);
            fills.Add(style.fill);
            fonts.Add(style.font);

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
            if (index >= 0)
                style = styles[index];
            else
            {
                styles.Add(style);
                if (!alignments.Contains(style.alignment)) alignments.Add(style.alignment);
                if (!borders.Contains(style.border)) borders.Add(style.border);
                if (!fills.Contains(style.fill)) fills.Add(style.fill);
                if (!fonts.Contains(style.font)) fonts.Add(style.font);
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
                if (fontNames.ContainsKey(style.font.name)) fontNames[style.font.name] += style.count;
                else fontNames.Add(style.font.name, style.count);
                if (colors.ContainsKey(style.font.color)) colors[style.font.color] += style.count;
                else colors.Add(style.font.color, style.count);
                if (colors.ContainsKey(style.fill.foreColor)) colors[style.fill.foreColor] += style.count;
                else colors.Add(style.fill.foreColor, style.count);
                if (colors.ContainsKey(style.fill.backColor)) colors[style.fill.backColor] += style.count;
                else colors.Add(style.fill.backColor, style.count);
            }
            var fontSorted = (from entry in fontNames orderby entry.Value descending select entry.Key).ToList();
            var colorSorted = (from entry in colors orderby entry.Value descending select entry.Key).ToList();
            foreach (var style in styles)
            {
                style.font.nameIndex = fontSorted.IndexOf(style.font.name);
                style.font.colorIndex = colorSorted.IndexOf(style.font.color);
                style.fill.foreIndex = colorSorted.IndexOf(style.fill.foreColor);
                style.fill.backIndex = colorSorted.IndexOf(style.fill.backColor);
            }
        }

        public override string ToString()
            => $"[R{top}C{left}:R{bottom}C{right}] Styles{styles.Count}";
    }
}
