using System.Text.RegularExpressions;

namespace HMarkupClassifier.SheetParser
{
    class Formula
    {
        public static Regex A1Regex = new Regex(@"(?<Col>[A-Z]{1,3})(?<Row>\d+)");
        public static Regex A1A1Regex = new Regex(@"(?<Left>[A-Z]{1,3})(?<Top>\d+):(?<Right>[A-Z]{1,3})(?<Bottom>\d+)");
    }
}
