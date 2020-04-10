using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
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
        }

        public static string CSVTitle
            = $"";//TODO

        public override string ToString()
        {
            return base.ToString();//TODO
        }
    }
}
