using System;

using ClosedXML.Excel;

namespace HMarkupClassifier.SheetParser
{
    class CellFeature
    {
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
        // Cell Style
        public Style style;
        public byte merged;
        public double width, height;
        public double widthRatio, heightRatio;

        // Value Feature
        public byte dataType;
        public byte empty;
        public int valueLen;
        // HasFormula, HasArrayFormula, HasComment, HasHyperlink
        public byte hasFormula;
        public byte hasArrayFormula;
        public byte hasHyperlink;
        public byte hasComment;
        // Content
        // TODO:

        // Position Feature
        public int col, row;
        public float leftRatio, topRatio, rightRatio, bottomRatio;
        public Position position = new Position();

        public CellFeature(IXLCell cell, SheetInfo info)
        {
            // Spatial
            col = cell.Address.ColumnNumber;
            row = cell.Address.RowNumber;
            leftRatio = (col - info.left) / (float)(info.NumCol);
            topRatio = (row - info.top) / (float)(info.NumRow);
            rightRatio = (info.right - col) / (float)(info.NumCol);
            bottomRatio = (info.bottom - row) / (float)(info.NumRow);

            // Cell Style
            style = info.GetStyle(cell.Style);
            merged = (byte)(cell.IsMerged() ? 1 : 0);
            width = cell.WorksheetColumn().Width;
            height = cell.WorksheetRow().Height;
            widthRatio = width / info.width;
            heightRatio = height / info.height;

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
            valueLen = value.Length;
        }

        public void SetPosition(int sideIndex, CellFeature cell) => position.SetPosition(sideIndex, this, cell);

        public static string csvTitle = $"col,row," +
               $"{Style.csvTitle},merged,width,height,widthRatio,heightRatio," +
               $"dataType,valueLen,empty,hasArrayFormula,hasComment,hasHyperlink,hasFormula," +
               $"leftRatio,topRatio,rightRatio,bottomRatio,{Position.csvTitle}";

        public string CSVString()
            => $"{col},{row}," +
               $"{style.CSVString()},{merged},{width},{height},{widthRatio},{heightRatio}," +
               $"{dataType},{valueLen},{empty},{hasArrayFormula},{hasComment},{hasHyperlink},{hasFormula}," +
               $"{leftRatio},{topRatio},{rightRatio},{bottomRatio},{position.CSVString()}";
    }
}
