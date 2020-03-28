

namespace HMarkupClassifier.SheetParser
{
    class Position
    {
        public Side[] sides = new Side[4];

        public Position()
        {
            for (int i = 0; i < 4; i++)
            {
                sides[i] = new Side();
            }
        }

        public void SetPosition(int sideIndex, CellFeature self, CellFeature cell)
        {
            var side = sides[sideIndex];
            if (self.style.Equals(cell.style)) side.style = 1;
            if (self.style.alignment.Equals(cell.style.alignment)) side.alignment = 1;
            if (self.style.fill.Equals(cell.style.fill)) side.fill = 1;
            if (self.style.font.Equals(cell.style.font)) side.font = 1;
            if (self.width == cell.width) side.width = 1;
            if (self.height == cell.height) side.height = 1;
            side.dataType = cell.dataType;
            side.empty = cell.empty;
            side.words = cell.words;
            side.referenced = cell.isReferenced;
        }

        public static string csvTitle = $"{Side.csvTitle},{Side.csvTitle},{Side.csvTitle},{Side.csvTitle}";

        public string CSVString()
            => $"{sides[0].CSVString()},{sides[1].CSVString()},{sides[2].CSVString()},{sides[3].CSVString()}";
    }

    class Side
    {
        public byte style = 0;
        public byte alignment = 0;
        public byte fill = 0;
        public byte font = 0;
        public byte width = 0;
        public byte height = 0;

        public byte dataType = 0;
        public byte empty = 1;
        public int words = 0;
        public byte referenced = 0;

        public static string csvTitle = "StyleDiff,NeighborDataType,NeighborEmpty,NeighborWords,NeighborRfd";

        public string CSVString()
            => $"{style + alignment + fill + font + width + height},{dataType},{empty},{words},{referenced}";
    }
}
