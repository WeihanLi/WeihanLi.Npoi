using NPOI.SS.UserModel;
using System;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi.Configurations
{
    public sealed class SheetConfiguration
    {
        private int _startRowIndex = 1;
        private string _sheetName = "Sheet0";
        private int _sheetIndex;
        private Func<IRow, bool> _rowFilter;
        private Func<ICell, bool> _cellFilter;

        public SheetConfiguration()
        {
            _rowFilter = row => true;
            _cellFilter = cell => true;
        }

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
        /// EndRowIndex, included
        /// </summary>
        public int? EndRowIndex { get; set; }

        /// <summary>
        /// enable auto column width
        /// </summary>
        public bool AutoColumnWidthEnabled { get; set; }

        /// <summary>
        /// Cell Filter
        /// </summary>
        public Func<ICell, bool> CellFilter
        {
            get => _cellFilter;
            set => _cellFilter = value ?? (cell => true);
        }

        /// <summary>
        /// Row Filter
        /// </summary>
        public Func<IRow, bool> RowFilter
        {
            get => _rowFilter;
            set => _rowFilter = value ?? (row => true);
        }
    }
}
