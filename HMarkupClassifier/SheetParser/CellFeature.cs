using System;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class CellFeature
    {
        // Cell Style
        public Style style;
        public byte merged;
        public double width, height, fontWidth, fontHeight;
        public double widthRatio, heightRatio;

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
        private static readonly Regex beginNumRegex = new Regex(@"^(?:[\W]*\d+)[^\d]+", RegexOptions.Compiled);
        public byte beginNumber;
        private static readonly Regex beginSpecialRegex = new Regex(@"^[\W]", RegexOptions.Compiled);
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

        // Position Feature
        public int col, row;
        public float leftRatio, topRatio;
        public Position neighbors = new Position();
        public Position nonEmptyNeighbors = new Position();
        public int leftDistance, topDistance, rightDistance, bottomDistance;

        public CellFeature(IXLCell cell, SheetInfo info)
        {
            // Position
            col = cell.Address.ColumnNumber;
            row = cell.Address.RowNumber;
            leftRatio = (col - info.left) / (float)(info.NumCol);
            topRatio = (row - info.top) / (float)(info.NumRow);
            leftDistance = col - info.left;
            topDistance = row - info.top;
            rightDistance = info.right - col;
            bottomDistance = info.bottom - row;

            // Cell Style
            style = info.GetStyle(cell.Style);
            merged = (byte)(cell.IsMerged() ? 1 : 0);
            width = cell.WorksheetColumn().Width;
            fontWidth = width / style.font.size;
            height = cell.WorksheetRow().Height;
            fontHeight = height / style.font.size;
            widthRatio = width / info.width;
            heightRatio = height / info.height;

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
                beginNumber = (byte)(beginNumRegex.IsMatch(value) ? 1 : 0);
                beginSpecial = (byte)(beginSpecialRegex.IsMatch(value) ? 1 : 0);
                symbols = (byte)(symbolsRegex.IsMatch(value) ? 1 : 0);
                punctuation = (byte)(punctuationRegex.IsMatch(value) ? 1 : 0);
                allUpper = (byte)(hasLowerRegex.IsMatch(value) ? 0 : 1);
                allNumber = (byte)(allNumberRegex.IsMatch(value) ? 1 : 0);
                if (value.Length != 0)
                    textWidth = fontWidth / value.Length;
            }
        }

        public void SetNeighbors(int sideIndex, CellFeature cell) => neighbors.SetPosition(sideIndex, this, cell);

        public void SetNonEmptyNeighbors(int sideIndex, CellFeature cell) => nonEmptyNeighbors.SetPosition(sideIndex, this, cell);

        public static string csvTitle = $"col,row," +
               $"{Style.csvTitle},Merged,FontWidth,FontHeight,WidthRatio,HeightRatio," +
               $"DataType,Words,empty," +
               $"BeginNumber,BeginSpecial,Symbols,Punctuation,AllNumber,AllUpper,TextWidth" +
               $"IsReferenced,ArrayFormula,Comment,Hyperlink,Formula," +
               $"LeftDistance,TopDistance,RightDistance,BottomDistance," +
               $"LeftRatio,TopRatio,{Position.csvTitle},{Position.csvTitle}";

        public string CSVString()
            => $"{col},{row}," +
               $"{style.CSVString()},{merged},{fontWidth},{fontHeight},{widthRatio},{heightRatio}," +
               $"{dataType},{words},{empty}," +
               $"{beginNumber},{beginSpecial},{symbols},{punctuation},{allNumber},{allUpper},{textWidth}" +
               $"{isReferenced},{hasArrayFormula},{hasComment},{hasHyperlink},{hasFormula}," +
               $"{leftDistance},{topDistance},{rightDistance},{bottomDistance}," +
               $"{leftRatio},{topRatio},{neighbors.CSVString()},{nonEmptyNeighbors.CSVString()}";
    }
}
