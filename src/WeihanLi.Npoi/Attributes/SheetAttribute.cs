using System;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class SheetAttribute : Attribute
    {
        public int SheetIndex { get => SheetSetting.SheetIndex; set => SheetSetting.SheetIndex = value; }

        public string SheetName { get => SheetSetting.SheetName; set => SheetSetting.SheetName = value; }

        public int StartRowIndex { get => SheetSetting.StartRowIndex; set => SheetSetting.StartRowIndex = value; }

        public int HeaderRowIndex => SheetSetting.HeaderRowIndex;

        public int? EndRowIndex { get => SheetSetting.EndRowIndex; set => SheetSetting.EndRowIndex = value; }

        private int _startColumnIndex;
        private int? _endColumnIndex;

        /// <summary>
        /// StartColumnIndex
        /// Start Column Index when import
        /// </summary>
        public int StartColumnIndex
        {
            get => _startColumnIndex;
            set
            {
                if (value >= 0)
                {
                    if (_endColumnIndex.HasValue && _endColumnIndex.Value >= value)
                    {
                        SheetSetting.CellFilter = cell =>
                            cell.ColumnIndex >= value && cell.ColumnIndex <= _endColumnIndex.Value
                            ;
                    }
                    else
                    {
                        SheetSetting.CellFilter = cell => cell.ColumnIndex >= value;
                    }
                    _startColumnIndex = value;
                }
            }
        }

        /// <summary>
        /// EndColumnIndex
        /// End Column Index when import
        /// </summary>
        public int? EndColumnIndex
        {
            get => _endColumnIndex;
            set
            {
                if (value.HasValue)
                {
                    if (value.Value >= _startColumnIndex)
                    {
                        SheetSetting.CellFilter = cell =>
                                cell.ColumnIndex >= value && cell.ColumnIndex <= _endColumnIndex.Value
                            ;
                        _endColumnIndex = value;
                    }
                }
                else
                {
                    SheetSetting.CellFilter = cell => cell.ColumnIndex >= _startColumnIndex;
                    _endColumnIndex = null;
                }
            }
        }

        public bool AutoColumnWidthEnabled
        {
            get => SheetSetting.AutoColumnWidthEnabled;
            set => SheetSetting.AutoColumnWidthEnabled = value;
        }

        internal SheetConfiguration SheetSetting { get; }

        public SheetAttribute() => SheetSetting = new SheetConfiguration();
    }
}
