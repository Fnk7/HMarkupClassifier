﻿using System;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class CellABD
    {
        // Cell Style
        public XStyle style;
        public int merged;
        public double width, height;

        // Value Feature
        public byte dataType;
        public byte empty;
        public int words = 0;
        // HasFormula, HasArrayFormula, HasComment, HasHyperlink
        public byte isReferenced = 0;
        public byte hasFormula;
        public byte hasArrayFormula;
        public byte hasHyperlink;
        public byte hasComment;
        // Content
        private static readonly Regex likeYearRegex = new Regex(@"^(?:19)|(?:20)\d{2}$");
        private byte likeYear;
        private static readonly Regex beginNumRegex = new Regex(@"^(?:[\W]*\d+)[^\d]+", RegexOptions.Compiled);
        public byte beginNumber;
        private static readonly Regex beginSpecialRegex = new Regex(@"^[^\p{L}]", RegexOptions.Compiled);
        public byte beginSpecial;
        private static readonly Regex symbolsRegex = new Regex(@"[\p{S}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        public byte symbols;
        private static readonly Regex punctuationRegex = new Regex(@"[\p{P}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        public byte punctuation;
        private static readonly Regex hasLowerRegex = new Regex(@"[\p{Ll}]", RegexOptions.Compiled);
        public byte allUpper;
        private static readonly Regex allNumberRegex = new Regex(@"^(?:[^\p{L}]*\d+[^\p{L}]*)+$", RegexOptions.Compiled);
        public byte allNumber;

        // Position Feature
        public int col, row;
        public float leftRatio, topRatio;
        public Position neighbors = new Position();
        public Position nonEmptyNeighbors = new Position();

        public CellABD(IXLCell cell, SheetInfoABD info)
        {
            // Position
            col = cell.Address.ColumnNumber;
            row = cell.Address.RowNumber;
            leftRatio = (col - info.left) / (float)(info.NumCol);
            topRatio = (row - info.top) / (float)(info.NumRow);

            // Cell Style
            style = info.GetStyle(cell.Style);
            merged = (byte)(cell.IsMerged() ? 1 : 0);
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;

            // Formula
            if (cell.HasFormula)
                info.formulas.Add(cell.FormulaA1);

            // Value
            dataType = (byte)cell.DataType;
            string value;
            try
            {
                IXLCell temp = cell;
                if (cell.IsMerged()) temp = cell.MergedRange().FirstCell();
                hasFormula = (byte)(temp.HasFormula ? 1 : 0);
                hasArrayFormula = (byte)(temp.HasArrayFormula ? 1 : 0);
                hasHyperlink = (byte)(temp.HasHyperlink ? 1 : 0);
                hasComment = (byte)(temp.HasComment ? 1 : 0);
                value = temp.GetString();
            }
            catch (Exception ex) { throw new ParseFailException("Can't Get Value of Cell!", ex); }
            empty = (byte)(value.Length == 0 ? 1 : 0);
            if (empty == 0)
            {
                words = Regex.Split(value, @"\W+").Length;
                likeYear = (byte)(likeYearRegex.IsMatch(value.Trim()) ? 1 : 0);
                beginNumber = (byte)(beginNumRegex.IsMatch(value) ? 1 : 0);
                beginSpecial = (byte)(beginSpecialRegex.IsMatch(value) ? 1 : 0);
                symbols = (byte)(symbolsRegex.IsMatch(value) ? 1 : 0);
                punctuation = (byte)(punctuationRegex.IsMatch(value) ? 1 : 0);
                allUpper = (byte)(hasLowerRegex.IsMatch(value) ? 0 : 1);
                allNumber = (byte)(allNumberRegex.IsMatch(value) ? 1 : 0);
            }
        }

        public void SetNeighbors(int sideIndex, CellABD cell)
        {
            neighbors.SetPosition(sideIndex, this, cell);
        }
        public void SetNonEmptyNeighbors(int sideIndex, CellABD cell) => nonEmptyNeighbors.SetPosition(sideIndex, this, cell);

        public static string csvTitle = $"col,row,empty," +
               $"{XStyle.CSVTitle},merged," +
               $"datatype,words," +
               $"likeyear,beginnum,beginspecial,symbolspunctuation,allnum,allupper," +
               $"isreferenced,arrayformula,comment,hyperlink,formula," +
               $"LeftRatio,TopRatio,{Position.CSVTitle},{Position.CSVTitle}";

        public string CSVString()
            => $"{col},{row},{empty}," +
               $"{style},{merged}," +
               $"{dataType},{words}," +
               $"{likeYear},{beginNumber},{beginSpecial},{symbols},{punctuation},{allNumber},{allUpper}," +
               $"{isReferenced},{hasArrayFormula},{hasComment},{hasHyperlink},{hasFormula}," +
               $"{leftRatio},{topRatio},{neighbors.CSVString()},{nonEmptyNeighbors.CSVString()}";
    }
}