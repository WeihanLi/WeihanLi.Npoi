using WeihanLi.Extensions;

namespace WeihanLi.Document.Settings.Excel
{
    internal sealed class SheetSetting
    {
        private int _startRowIndex = 1;
        private string _sheetName = "Sheet0";
        private int _sheetIndex;

        /// <summary>
        /// SheetIndex
        /// </summary>
        public int SheetIndex
        {
            get => _sheetIndex;
            set
            {
                if (value >= 0)
                {
                    _sheetIndex = value;
                }
            }
        }

        /// <summary>
        /// SheetName
        /// </summary>
        public string SheetName
        {
            get => _sheetName;
            set
            {
                if (value.IsNotNullOrWhiteSpace())
                {
                    _sheetName = value;
                }
            }
        }

        /// <summary>
        /// StartRowIndex
        /// </summary>
        public int StartRowIndex
        {
            get => _startRowIndex;
            set
            {
                if (value >= 0)
                {
                    _startRowIndex = value;
                }
            }
        }

        /// <summary>
        /// HeaderRowIndex
        /// </summary>
        public int HeaderRowIndex => StartRowIndex - 1;

        /// <summary>
        /// enable auto column width
        /// </summary>
        public bool AutoColumnWidthEnabled { get; set; }
    }
}
