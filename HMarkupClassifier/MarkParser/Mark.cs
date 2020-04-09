namespace HMarkupClassifier.MarkParser
{
    class Mark
    {
        public int type, left, top, right, bottom;

        public Mark(int type, int left, int top, int right, int bottom)
        {
            this.type = type;
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        public bool ContainCell(int col, int row)
            => col >= left && col <= right && row >= top && row <= bottom;

        public override string ToString()
            => $"[Mk{type}R{top}C{left}:R{bottom}C{left}]";
    }
}
