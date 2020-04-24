using System.Linq;
using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class XContent
    {
        // Content
        public int Length;
        public int Words;
        public static readonly Regex YearRegex = new Regex(@"(?:19)|(?:20)\d{2}");
        public int Year;
        public static readonly Regex BeginRegex = new Regex(@"^[\p{S}\p{P}]", RegexOptions.Compiled);
        public int Begin;
        public static readonly Regex SpaceRegex = new Regex(@"^[\s]+");
        public int Space;
        public static readonly Regex SymbolRegex = new Regex(@"[\p{S}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        public int Symbol;
        public static readonly Regex PunctuationRegex = new Regex(@"[\p{P}]", RegexOptions.Compiled & RegexOptions.CultureInvariant);
        public int Punctuation;
        public static string[] MetaWordList =
            {"revised", "remark", "note", "table", "source", "data"};
        public int MetaWord;
        public static string[] StatisticWordList =
            {"total", "mean", "avg", "average", "percent", "amount", "min", "max", "std", "all", "count"};
        public int StatisticWord;
        public static string[] DateWordList =
            {"Spr", "Sum", "Fall", "Aut", "Win",
             "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"};
        public int DateWord;
        public static string[] HeaderWordList =
            {"name", "year", "country", "mail", "ii", "iii", "iv", "id", "no."};
        public int HeaderWord;
        public static string[] DataWordList =
            {"n/a", "~", "-", "a", "b", "c", "d"};
        public int DataWord;
        public XContent()
        {
            Length = 0;
            Words = 0;
            Year = 0;
            Begin = 0;
            Space = 0;
            Symbol = 0;
            Punctuation = 0;
            MetaWord = 0;
            StatisticWord = 0;
            DateWord = 0;
            HeaderWord = 0;
            DataWord = 0;
        }

        public XContent(string content)
        {
            Length = content.Length;
            Words = Regex.Split(content, @"\W+").Length;
            string trimContent = content.Trim();
            string ignoreCaseContent = trimContent.ToLower();
            Year = YearRegex.IsMatch(trimContent) ? 1 : 0;
            Begin = BeginRegex.IsMatch(trimContent) ? 1 : 0;
            var spaceMatch = SpaceRegex.Match(content);
            if (spaceMatch.Success)
            {
                Space = spaceMatch.Value.Length;
            }
            else
            {
                Space = 0;
            }
            Symbol = SymbolRegex.IsMatch(trimContent) ? 1 : 0;
            Punctuation = PunctuationRegex.IsMatch(trimContent) ? 1 : 0;
            MetaWord = MetaWordList
                .Any(word => ignoreCaseContent.Contains(word)) ? 1 : 0;
            StatisticWord = StatisticWordList
                .Any(word => ignoreCaseContent.Contains(word)) ? 1 : 0;
            DateWord = DateWordList
                .Any(word => trimContent.Contains(word)) ? 1 : 0;
            HeaderWord = HeaderWordList
                .Any(word => ignoreCaseContent.Contains(word)) ? 1 : 0;
            DataWord = DataWordList
                .Any(word => ignoreCaseContent == word) ? 1 : 0;
        }

        public int Similar(XContent content)
        {
            if (Length == 0 || content.Length == 0)
                return 0;
            int similarity = 0;
            if (Words == content.Words)
                similarity++;
            if (Year == content.Year)
                similarity += Year;
            if (Begin == content.Begin)
                similarity += Begin;
            if (Space != 0 && content.Space != 0)
                similarity ++;
            if (Symbol == content.Symbol)
                similarity += Symbol;
            if (Punctuation == content.Punctuation)
                similarity += Punctuation;
            if (MetaWord == content.MetaWord)
                similarity += MetaWord;
            if (StatisticWord == content.StatisticWord)
                similarity += StatisticWord;
            if (DateWord == content.DateWord)
                similarity += DateWord;
            if (HeaderWord == content.HeaderWord)
                similarity += HeaderWord;
            if (DataWord == content.DataWord)
                similarity += DataWord;
            return similarity;
        }


        public static readonly string p = "ctnt-";

        public static string CSVTitle
            = $"{p}length,{p}words," +
              $"{p}year,{p}begin,{p}space," +
              $"{p}symbol,{p}punctuation," +
              $"{p}meta,{p}statistic,{p}date," +
              $"{p}header,{p}data";

        public override string ToString()
        {
            return $"{Length},{Words}," +
                   $"{Year},{Begin},{Space}," +
                   $"{Symbol},{Punctuation}," +
                   $"{MetaWord},{StatisticWord},{DateWord}," +
                   $"{HeaderWord},{DataWord}";
        }
    }
}
