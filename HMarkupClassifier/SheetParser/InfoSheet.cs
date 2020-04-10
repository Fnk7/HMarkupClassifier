using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using HMarkupClassifier.SheetParser.Styles;

namespace HMarkupClassifier.SheetParser
{
    class InfoSheet
    {
        public int MaxStyles = 100;
        public Dictionary<XStyle, XStyle> styles = new Dictionary<XStyle, XStyle>();
        public List<string> FontNames = new List<string>();

        public XFormat GetXFormat(IXLCell cell)
        {
            XFormat xFormat = new XFormat(cell);
            XStyle xStyle = new XStyle(cell.Style);
            var FontName = cell.Style.Font.FontName;
            if (!FontNames.Contains(FontName))
                FontNames.Add(FontName);
            xStyle.Font.NameIndex = FontNames.IndexOf(FontName);
            if (styles.ContainsKey(xFormat.Style))
                xStyle = styles[xStyle];
            else
                styles.Add(xStyle, xStyle);
            xStyle.count++;
            xFormat.Style = xStyle;
            return xFormat;
        }

        public XContent GetXContent(IXLCell cell)
        {
            if (cell.IsEmpty())
                return new XContent();
            try
            {
                string content = cell.GetString();
                return new XContent(content);
            }
            catch (Exception ex)
            {
                throw new ParseFailException("Parse Content Fail!", ex);
            }
        }

        public List<string> FormulaA1s = new List<string>();

        public XFormula GetXFormula(IXLCell cell)
        {
            XFormula xFormula = new XFormula();
            if (cell.HasFormula)
                FormulaA1s.Add(cell.FormulaA1);
            xFormula.HasFormula = cell.HasFormula ? 1 : 0;
            xFormula.HasArrayFormula = cell.HasArrayFormula ? 1 : 0;
            return xFormula;
        }

        public void SetFormulaReferenced(Dictionary<(int, int), XCell> cells)
        {
            foreach (string Formula in FormulaA1s)
            {
                foreach (Match match in XFormula.A1Regex.Matches(Formula.ToUpper()))
                {
                    int left = Utils.ParseColumn(match.Groups["Left"].Value);
                    int top = Convert.ToInt32(match.Groups["Top"].Value);
                    if (match.Groups["Right"].Value != string.Empty)
                    {
                        int right = Utils.ParseColumn(match.Groups["Right"].Value);
                        int bottom = Convert.ToInt32(match.Groups["Bottom"].Value);
                        for (int row = top; row <= top; row++)
                        {
                            for (int col = left; col <= right; col++)
                            {
                                if (cells.ContainsKey((row, col)))
                                {
                                    cells[(row, col)].Formula.IsReferenced += 1;
                                }
                            }
                        }
                    }
                    else if(cells.ContainsKey((top, left)))
                    {
                        cells[(top, left)].Formula.IsReferenced += 1;
                    }
                }
            }
        }

        public void SetOrderedIndex()
        {
            Dictionary<int, int> FontNameCnt = new Dictionary<int, int>();
            Dictionary<int, int> FontColorCnt = new Dictionary<int, int>();
            Dictionary<int, int> PtnColorCnt = new Dictionary<int, int>();
            Dictionary<int, int> BckColorCnt = new Dictionary<int, int>();
            foreach (var style in styles.Keys)
            {
                if (!FontNameCnt.ContainsKey(style.Font.NameIndex))
                    FontNameCnt.Add(style.Font.NameIndex, 0);
                if (!FontColorCnt.ContainsKey(style.Font.Color))
                    FontColorCnt.Add(style.Font.Color, 0);
                FontNameCnt[style.Font.NameIndex] += style.count;
                FontColorCnt[style.Font.Color] += style.count;
                if (!PtnColorCnt.ContainsKey(style.Fill.PtnColor))
                    PtnColorCnt.Add(style.Fill.PtnColor, 0);
                if (!BckColorCnt.ContainsKey(style.Fill.BckColor))
                    BckColorCnt.Add(style.Fill.BckColor, 0);
                PtnColorCnt[style.Fill.PtnColor] += style.count;
                PtnColorCnt[style.Fill.BckColor] += style.count;
            }
            var OrderedFontName = (from x in FontNameCnt orderby x.Value descending select x.Key).ToList();
            var OrderedFontColor = (from x in FontColorCnt orderby x.Value descending select x.Key).ToList();
            var OrderedPtnColor = (from x in PtnColorCnt orderby x.Value descending select x.Key).ToList();
            var OrderedBckColor = (from x in BckColorCnt orderby x.Value descending select x.Key).ToList();
            foreach (var style in styles.Keys)
            {
                style.Font.NameIndex = OrderedFontName.IndexOf(style.Font.NameIndex);
                style.Font.Color = OrderedFontName.IndexOf(style.Font.Color);
                style.Fill.PtnColor = OrderedPtnColor.IndexOf(style.Fill.PtnColor);
                style.Fill.PtnColor = OrderedBckColor.IndexOf(style.Fill.BckColor);
            }
        }
    }
}
