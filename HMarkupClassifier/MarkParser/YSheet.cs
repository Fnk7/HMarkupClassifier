using System.Collections.Generic;
using System.Linq;

namespace HMarkupClassifier.MarkParser
{
    class YSheet
    {
        public List<Mark> marks;
        public YSheet(List<Mark> marks) => this.marks = marks;

        public int GetCellType(int col, int row)
        {
            if (marks.FirstOrDefault(m => m.ContainCell(col, row)) is Mark mark)
                return mark.type;
            return 3;
        }

        public override string ToString()
            => string.Concat(marks);
    }
}
