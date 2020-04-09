using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class Content
    {
        // Value Feature
        public int dataType;
        public int notEmpty ;
        // HasFormula, HasArrayFormula, HasComment, HasHyperlink
        public int referenced = 0;
        public int hasFormula = 0;
        public int hasArrayFormula = 0;
        public int hasHyperlink = 0;
        public int hasComment = 0;

        // Content
        public int Words = 0;
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
        public double textWidth;

        public Content(IXLCell cell, Statistic statistic)
        {
            string value = string.Empty;
            try
            {
                if (cell.IsMerged())
                    value = cell.MergedRange().FirstCell().GetString();
                else
                    value = cell.GetString();
            }
            catch (Exception ex) { throw new ParseFailException("Can't Get Value of Cell!", ex); }
            dataType = (int)cell.DataType;
            notEmpty = value.Length == 0 ? 0 : 1;

        }

        public void SetReferenced() => referenced = 1;
    }
}
