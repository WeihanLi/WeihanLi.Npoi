namespace WeihanLi.Npoi.Settings
{
    internal class SheetSetting
    {
        /// <summary>
        /// SheetIndex
        /// </summary>
        public int SheetIndex { get; internal set; }

        /// <summary>
        /// SheetName
        /// </summary>
        public string SheetName { get; internal set; }

        /// <summary>
        /// StartRowIndex
        /// </summary>
        public int StartRowIndex { get; internal set; } = 1;

        /// <summary>
        /// HeaderRowIndex
        /// </summary>
        public int HeaderRowIndex => StartRowIndex - 1;
    }
}
