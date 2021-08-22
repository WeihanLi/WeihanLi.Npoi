namespace WeihanLi.Npoi.Settings
{
    internal sealed class FreezeSetting
    {
        public FreezeSetting(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
        {
        }

        public FreezeSetting(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            ColSplit = colSplit;
            RowSplit = rowSplit;
            LeftMostColumn = leftmostColumn;
            TopRow = topRow;
        }

        /// <summary>
        ///     horizontal position of split
        /// </summary>
        public int ColSplit { get; }

        /// <summary>
        ///     Vertical position of split
        /// </summary>
        public int RowSplit { get; }

        /// <summary>
        ///     Top row visible in bottom pane
        /// </summary>
        public int LeftMostColumn { get; }

        /// <summary>
        ///     Left column visible in right pane
        /// </summary>
        public int TopRow { get; }
    }
}
