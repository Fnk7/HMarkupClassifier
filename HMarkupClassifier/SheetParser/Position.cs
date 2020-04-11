

namespace HMarkupClassifier.SheetParser
{
    class Position
    {
        public Side[] sides = new Side[4];

        public Position()
        {
            for (int i = 0; i < 4; i++)
                sides[i] = new Side();
        }

        //public void SetPosition(int sideIndex, CellABD self, CellABD cell)
        //{
        //    var side = sides[sideIndex];
        //    if (self.style.Equals(cell.style)) side.style += 1;
        //    if (self.style.Alignment.Equals(cell.style.Alignment)) side.style += 1;
        //    if (self.style.Fill.Equals(cell.style.Fill)) side.style += 1;
        //    if (self.style.Font.Equals(cell.style.Font)) side.style += 1;
        //    side.datatype = cell.dataType;
        //    side.refered = cell.isReferenced;
        //    side.formula = cell.hasFormula;
        //    side.words = cell.words;
        //}

        public static string CSVTitle
            = $"{Side.CSVTitle(0)},{Side.CSVTitle(1)},{Side.CSVTitle(2)},{Side.CSVTitle(3)}";

        public string CSVString()
            => $"{sides[0].CSVString()},{sides[1].CSVString()},{sides[2].CSVString()},{sides[3].CSVString()}";
    }

    class Side
    {
        public byte style = 0;
        public byte datatype = 0;
        public byte refered = 0;
        public byte formula = 0;
        public int words = 0;

        public static string[] names = { "s-left", "s-top", "s-right", "s-bottom" };
        public static string CSVTitle(int index)
            => $"{names[index]}-style," +
               $"{names[index]}-datatype," +
               $"{names[index]}-formula," +
               $"{names[index]}-refered," +
               $"{names[index]}-words";

        public string CSVString()
            => $"{style},{datatype},{refered},{formula},{words}";
    }
}
