using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class XFormula
    {
        public static Regex A1Regex = new Regex(@"\$?(?<Left>[A-Z]{1,3})\$?(?<Top>\d+)(?::\$?(?<Right>[A-Z]{1,3})\$?(?<Bottom>\d+))?");

        public int HasFormula = 0; 
        public int HasArrayFormula = 0;
        public int IsReferenced = 0;

        public int Similar(XFormula xFormula)
        {
            int similarity = 0;
            if (HasFormula == xFormula.HasFormula)
                similarity++;
            if (IsReferenced != 0 && xFormula.IsReferenced != 0)
            {
                similarity++;
                if (IsReferenced == xFormula.IsReferenced)
                    similarity++;
            }
            return similarity;
        }

        public static string CSVTitle
            = $"fml-formula,fml-arrayformula,fml-referenced";

        public override string ToString()
        {
            return $"{HasFormula},{HasArrayFormula},{IsReferenced}";
        }
    }
}
