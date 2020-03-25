using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class CellFeature
    {
        private static readonly Regex capitalRegex = new Regex(@"^[^a-z]+$", RegexOptions.Compiled);
        private static readonly Regex nonAlphabetRegex = new Regex(@"^[^a-zA-Z]$");
        /*
         * Address
         * DataType
         * FormulaA1
         * HasFormula, HasArrayFormula
         * HasComment
         * HasHyperlink
         * Style
         * Value
         */
        #region flag
        private int _flag = 0;
        private static readonly int hasBorder = 1 << 0;
        private static readonly int isFill = 1 << 1;
        private static readonly int isEmpty = 1 << 2;
        private static readonly int allCapital = 1 << 3;
        private static readonly int beginWithSpecial = 1 << 4;
        private static readonly int nonAlphabet = 1 << 5;
        private static readonly int hasFormual = 1 << 6;
        private static readonly int hasArrayFormula = 1 << 7;
        private static readonly int hasHyperlink = 1 << 8;
        private static readonly int hasComment = 1 << 9;
        #endregion

        // HasFormula, HasArrayFormula, HasComment, HasHyperlink
        public bool HasFormula
        {
            get { return (_flag & hasFormual) != 0; }
            set { _flag = value ? (_flag | hasFormual) : (_flag & ~hasFormual); }
        }
        public bool HasArrayFormula
        {
            get { return (_flag & hasArrayFormula) != 0; }
            set { _flag = value ? (_flag | hasArrayFormula) : (_flag & ~hasArrayFormula); }
        }
        public bool HasHyperlink
        {
            get { return (_flag & hasHyperlink) != 0; }
            set { _flag = value ? (_flag | hasHyperlink) : (_flag & ~hasHyperlink); }
        }
        public bool HasComment
        {
            get { return (_flag & hasComment) != 0; }
            set { _flag = value ? (_flag | hasComment) : (_flag & ~hasComment); }
        }
        // Value


        // Cell
        private double width, height;
        private Style style;

        // Value
        public int valueLen;
        public int dataType;
        public bool IsEmpty
        {
            get { return (_flag & isEmpty) != 0; }
            set { _flag = value ? (_flag | isEmpty) : (_flag & ~isEmpty); }
        }
        public bool AllCapital
        {
            get { return (_flag & allCapital) != 0; }
            set { _flag = value ? (_flag | allCapital) : (_flag & ~allCapital); }
        }
        public bool BeginWithSpecial
        {
            get { return (_flag & beginWithSpecial) != 0; }
            set { _flag = value ? (_flag | beginWithSpecial) : (_flag & ~beginWithSpecial); }
        }
        public bool NonAlphabet
        {
            get { return (_flag & nonAlphabet) != 0; }
            set { _flag = value ? (_flag | nonAlphabet) : (_flag & ~nonAlphabet); }
        }

        // Spatial
        private int col, row;

        public CellFeature(IXLCell cell, List<Style> styles)
        {
            // Cell
            style = new Style(cell.Style);
            int index = styles.IndexOf(style);
            if (index >= 0) style = styles[index];
            else styles.Add(style);
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;

            // Value
            dataType = (int)cell.DataType;
            if (!cell.IsEmpty())
            {
                string value;
                try { value = cell.GetString(); }
                catch (Exception ex) { throw new ParseFailException("Can't Get Value of Cell!", ex); }
                valueLen = value.Length;
                if (capitalRegex.IsMatch(value)) AllCapital = true;
                if (nonAlphabetRegex.IsMatch(value)) NonAlphabet = true;
            }
            else
                IsEmpty = true;
            HasFormula = cell.HasFormula;
            HasArrayFormula = cell.HasArrayFormula;
            HasHyperlink = cell.HasHyperlink;
            HasComment = cell.HasComment;
            // Spatial
            col = cell.Address.ColumnNumber;
            row = cell.Address.RowNumber;
            // 
        }
    }
}
