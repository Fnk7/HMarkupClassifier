

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

        public void SetPosition(int sideIndex, CellFeature self, CellFeature cell)
        {
            var side = sides[sideIndex];
            if (self.style.Equals(cell.style)) side.styleDiff += 1;
            if (self.style.alignment.Equals(cell.style.alignment)) side.styleDiff += 1;
            if (self.style.fill.Equals(cell.style.fill)) side.styleDiff += 1;
            if (self.style.font.Equals(cell.style.font)) side.styleDiff += 1;
            side.sideLenRatio = sideIndex %  2 == 0 ?  cell.width / self.width : cell.height / self.height;
            side.dataType = cell.dataType;
            side.empty = cell.empty;
            side.referenced = cell.isReferenced;
            side.formula = cell.hasFormula;
            side.words = cell.words;
        }

        public static string csvTitle
            = $"{Side.CsvTitle(0)},{Side.CsvTitle(1)},{Side.CsvTitle(2)},{Side.CsvTitle(3)}";

        public string CSVString()
            => $"{sides[0].CSVString()},{sides[1].CSVString()},{sides[2].CSVString()},{sides[3].CSVString()}";
    }

    class Side
    {
        public byte styleDiff = 0;
        public double sideLenRatio;

        public byte dataType = 0;
        public byte empty = 1;
        public byte referenced = 0;
        public byte formula = 0;
        public int words = 0;

        private static string[] sideName = { "Left", "Top", "Right", "Bottom" };
        public static string CsvTitle(int sideIndex)
            => $"{sideName[sideIndex]}-StyleDiff," +
                $"{sideName[sideIndex]}-DataType," +
                $"{sideName[sideIndex]}-SideLenRatio," +
                $"{sideName[sideIndex]}-Formula," +
                $"{sideName[sideIndex]}-Referenced," +
                $"{sideName[sideIndex]}-Empty," +
                $"{sideName[sideIndex]}-Words";

        public string CSVString()
            => $"{styleDiff},{sideLenRatio},{dataType},{empty},{referenced},{formula},{words}";
    }
}
