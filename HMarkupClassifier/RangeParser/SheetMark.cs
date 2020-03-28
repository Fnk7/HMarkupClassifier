using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace HMarkupClassifier.RangeParser
{
    class SheetMark
    {
        private static readonly Regex sheetRegex = new Regex(@"\[(?<SheetName>[^\[\]]+)(?<Marks>(?:\[(?:(?:Tb)|(?:Mk-2))[^\[\]]+(?:\[[^\[\]]+\])*\])+)\]", RegexOptions.Compiled);
        private static readonly Regex markRegex = new Regex(@"\[Mk(?<Type>-?\d)\$(?<Left>[A-Z]+)\$(?<Top>\d+)(?::\$(?<Right>[A-Z]+)\$(?<Bottom>\d+))?\]", RegexOptions.Compiled);
        private List<Mark> marks;

        private SheetMark(List<Mark> marks) => this.marks = marks;

        public static Dictionary<string, SheetMark> ParseRange(string rangeFile)
        {
            string rangeString = File.ReadAllText(rangeFile);
            var sheetMatches = sheetRegex.Matches(rangeString);
            Dictionary<string, SheetMark> sheetMarks = new Dictionary<string, SheetMark>();
            foreach (Match sheetMatch in sheetMatches)
            {
                string sheetName = sheetMatch.Groups["SheetName"].Value;
                string marksString = sheetMatch.Groups["Marks"].Value;
                List<Mark> marks = new List<Mark>();
                var marksMatches = markRegex.Matches(marksString);
                foreach (Match marksMatch in marksMatches)
                {
                    int type, left, top, right, bottom;
                    type = Convert.ToInt32(marksMatch.Groups["Type"].Value);
                    left = Utils.ParseColumn(marksMatch.Groups["Left"].Value);
                    top = Convert.ToInt32(marksMatch.Groups["Top"].Value);
                    if (marksMatch.Groups["Right"].Value == "")
                    {
                        right = left;
                        bottom = top;
                    }
                    else
                    {
                        right = Utils.ParseColumn(marksMatch.Groups["Right"].Value);
                        bottom = Convert.ToInt32(marksMatch.Groups["Bottom"].Value);
                    }
                    marks.Add(new Mark(type, left, top, right, bottom));
                }
                if (marks.Count != 0) sheetMarks.Add(sheetName, new SheetMark(marks));
            }
            return sheetMarks;
        }

        public int GetCellType(int col, int row)
        {
            int type = 3;
            foreach (var mark in marks)
                if (type != mark.GetCellType(col, row)) return mark.GetCellType(col, row);
            return type;
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            foreach (var mark in marks)
                builder.Append(mark.ToString()).Append("\t");
            return builder.ToString();
        }
    }

    class Mark
    {
        private int type, left, top, right, bottom;

        public Mark(int type, int left, int top, int right, int bottom)
        {
            this.type = type;
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        public int GetCellType(int col, int row)
            => (col >= left && col <= right && row >= top && row <= bottom) ? type : 3;

        public override string ToString()
            => $"[Mk{type}R{top}C{left}:R{bottom}C{left}]";
    }
}
