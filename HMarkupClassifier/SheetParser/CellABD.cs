using System;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class CellABD
    {
        // Cell Style
        public double width, height;

        // Content
        //private static readonly Regex likeYearRegex = new Regex(@"^(?:19)|(?:20)\d{2}$");
        //private byte likeYear;
        //private static readonly Regex beginNumRegex = new Regex(@"^(?:[\W]*\d+)[^\d]+", RegexOptions.Compiled);
        //public byte beginNumber;
        //private static readonly Regex beginSpecialRegex = new Regex(@"^[^\p{L}]", RegexOptions.Compiled);
        //public byte beginSpecial;
        //private static readonly Regex symbolsRegex = new Regex(@"[\p{S}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        //public byte symbols;
        //private static readonly Regex punctuationRegex = new Regex(@"[\p{P}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        //public byte punctuation;
        //private static readonly Regex hasLowerRegex = new Regex(@"[\p{Ll}]", RegexOptions.Compiled);
        //public byte allUpper;
        //private static readonly Regex allNumberRegex = new Regex(@"^(?:[^\p{L}]*\d+[^\p{L}]*)+$", RegexOptions.Compiled);
        //public byte allNumber;

        // Position Feature
        public int col, row;
        public float leftRatio, topRatio;
        public Position neighbors = new Position();
        public Position nonEmptyNeighbors = new Position();

        public CellABD(IXLCell cell, SheetInfoABD info)
        {
            leftRatio = (col - info.left) / (float)(info.NumCol);
            topRatio = (row - info.top) / (float)(info.NumRow);

            // Cell Style
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;

            // Value
            //string value;
            //words = Regex.Split(value, @"\W+").Length;

        }

        //public void SetNeighbors(int sideIndex, CellABD cell)
        //{
        //    neighbors.SetPosition(sideIndex, this, cell);
        //}
        //public void SetNonEmptyNeighbors(int sideIndex, CellABD cell) => nonEmptyNeighbors.SetPosition(sideIndex, this, cell);

        public static string CSVTitle = $"col,row,empty," +
               $"LeftRatio,TopRatio,{Position.CSVTitle},{Position.CSVTitle}";

        public string CSVString()
            => $"{col},{row}," +
               $"{leftRatio},{topRatio}," +
               $"{neighbors.CSVString()},{nonEmptyNeighbors.CSVString()}";
    }
}
