using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class XFormula
    {
        public static Regex A1Regex = new Regex(@"\$?(?<Left>[A-Z]{1,3})\$?(?<Top>\d+)(?::\$?(?<Right>[A-Z]{1,3})\$?(?<Bottom>\d+))?");

        public int HasFormula = 0; 
        public int HasArrayFormula = 0;
        public int IsReferenced = 0;

        public static string CSVTitle
            = $"formula,arrayformula,referenced";

        public override string ToString()
        {
            return $"{HasFormula},{HasArrayFormula},{IsReferenced}";
        }
    }
}
