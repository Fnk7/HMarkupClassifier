using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.MarkParser
{
    static class MarkParserTool
    {
        public static readonly Regex markSheetRegex = new Regex(@"(?<Sheet>[^\[\]]+)(?<Marks>(?:\[(?:Tb|Mk),-?\d+,\d+,\d+,\d+,\d+\])+)", RegexOptions.Compiled);
        public static readonly Regex markMarksRegex = new Regex(@"\[Mk,(?<Type>-?\d),(?<Left>\d+),(?<Top>\d+),(?<Right>\d+),(?<Bottom>\d+)\]", RegexOptions.Compiled);

        private static readonly Regex rangeSheetRegex = new Regex(@"\[(?<Sheet>[^\[\]]+)(?<Marks>(?:\[(?:(?:Tb)|(?:Mk-2))[^\[\]]+(?:\[[^\[\]]+\])*\])+)\]", RegexOptions.Compiled);
        private static readonly Regex rangeMarksRegex = new Regex(@"\[Mk(?<Type>-?\d)\$(?<Left>[A-Z]+)\$(?<Top>\d+)(?::\$(?<Right>[A-Z]+)\$(?<Bottom>\d+))?\]", RegexOptions.Compiled);


        public static Dictionary<string, YSheet> Parse(string file)
        {
            if (file.EndsWith(".range"))
                return ParseRangeFile(file);
            else if (file.EndsWith(".mark"))
                return ParseMarkFile(file);
            else
                return new Dictionary<string, YSheet>();
        }

        public static Dictionary<string, YSheet> ParseRangeFile(string rangeFile)
        {
            Dictionary<string, YSheet> markSheets = new Dictionary<string, YSheet>();
            try
            {
                string rangeString = File.ReadAllText(rangeFile);
                foreach (Match sheetMatch in rangeSheetRegex.Matches(rangeString))
                {
                    string name = sheetMatch.Groups["Sheet"].Value;
                    List<Mark> marks = new List<Mark>();
                    foreach (Match marksMatch in rangeMarksRegex.Matches(sheetMatch.Groups["Marks"].Value))
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
                    if (marks.Count != 0) markSheets.Add(name, new YSheet(marks));
                }
            }
            catch (Exception) { }
            return markSheets;
        }

        public static Dictionary<string, YSheet> ParseMarkFile(string markFile)
        {
            Dictionary<string, YSheet> markSheets = new Dictionary<string, YSheet>();
            try
            {
                using (StreamReader markReader = new StreamReader(markFile))
                {
                    while (markReader.ReadLine() is string line)
                    {
                        Match sheetMatch = markSheetRegex.Match(line);
                        if (!sheetMatch.Success)
                            continue;
                        var name = sheetMatch.Groups["Sheet"].Value;
                        List<Mark> marks = new List<Mark>();
                        foreach (Match mark in markMarksRegex.Matches(sheetMatch.Groups["Marks"].Value))
                        {
                            int type = Convert.ToInt32(mark.Groups["Type"].Value);
                            int left = Convert.ToInt32(mark.Groups["Left"].Value);
                            int top = Convert.ToInt32(mark.Groups["Top"].Value);
                            int right = Convert.ToInt32(mark.Groups["Right"].Value);
                            int bottom = Convert.ToInt32(mark.Groups["Bottom"].Value);
                            marks.Add(new Mark(type, left, top, right, bottom));
                        }
                        if (marks.Count != 0) markSheets.Add(name, new YSheet(marks));
                    }
                }
            }
            catch (Exception) { }
            return markSheets;
        }
    }
}
