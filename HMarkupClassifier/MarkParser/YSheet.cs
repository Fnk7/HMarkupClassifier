using System.Collections.Generic;
using System.Linq;

namespace HMarkupClassifier.MarkParser
{
    class YSheet
    {
        public List<Mark> marks;
        public YSheet(List<Mark> marks) => this.marks = marks;

        public int GetCellType(int row, int col)
        {
            if (marks.FirstOrDefault(m => m.ContainCell(row, col)) is Mark mark)
                return mark.type;
            return 3;
        }

        public override string ToString()
            => string.Concat(marks);
    }
}
