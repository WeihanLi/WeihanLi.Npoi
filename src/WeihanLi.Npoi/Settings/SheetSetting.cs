namespace WeihanLi.Npoi.Settings
{
    internal class SheetSetting
    {
        /// <summary>
        /// SheetIndex
        /// </summary>
        internal int SheetIndex { get; set; }

        /// <summary>
        /// SheetName
        /// </summary>
        internal string SheetName { get; set; } = "Sheet0";

        /// <summary>
        /// StartRowIndex
        /// </summary>
        internal int StartRowIndex { get; set; } = 1;

        /// <summary>
        /// HeaderRowIndex
        /// </summary>
        internal int HeaderRowIndex => StartRowIndex - 1;
    }
}
