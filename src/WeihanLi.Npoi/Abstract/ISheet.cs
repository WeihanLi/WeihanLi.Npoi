namespace WeihanLi.Npoi.Abstract
{
    internal interface ISheet
    {
        /// <summary>
        /// FirstRowNum, 1 based rowNum
        /// 0 if no rows here
        /// </summary>
        int FirstRowNum { get; }

        /// <summary>
        /// lastRowIndex +1, 1 based rowNum
        /// 0 if no rows here
        /// </summary>
        int LastRowNum { get; }

        IRow? GetRow(int rowIndex);

        IRow CreateRow(int rowIndex);

        void SetColumnWidth(int columnIndex, int width);

        void AutoSizeColumn(int columnIndex);

        void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow);

        void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex);

        void ShiftRows(int startRow, int endRow, int n);

        IRow CopyRow(int sourceIndex, int targetIndex);

        void RemoveRow(IRow row);
    }
}
