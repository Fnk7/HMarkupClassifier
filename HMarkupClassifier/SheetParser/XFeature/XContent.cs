using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    // TODO 增加Content的Feature
    class XContent
    {
        // Content
        public int Words;
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

        public XContent()
        {
            Words = 0;
        }

        public XContent(string content)
        {
            Words = Regex.Split(content, @"\W+").Length;
            //likeYear = (byte)(likeYearRegex.IsMatch(value.Trim()) ? 1 : 0);
            //beginNumber = (byte)(beginNumRegex.IsMatch(value) ? 1 : 0);
            //beginSpecial = (byte)(beginSpecialRegex.IsMatch(value) ? 1 : 0);
            //symbols = (byte)(symbolsRegex.IsMatch(value) ? 1 : 0);
            //punctuation = (byte)(punctuationRegex.IsMatch(value) ? 1 : 0);
            //allUpper = (byte)(hasLowerRegex.IsMatch(value) ? 0 : 1);
            //allNumber = (byte)(allNumberRegex.IsMatch(value) ? 1 : 0);
        }

        public static string CSVTitle
            = $"words";

        public override string ToString()
        {
            return $"{Words}";
        }
    }
}
