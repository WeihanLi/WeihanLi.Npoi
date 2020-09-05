using System;
using WeihanLi.Document.Settings.Excel;

namespace WeihanLi.Document.Attributes.Excel
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class FilterAttribute : Attribute
    {
        internal FilterSetting FilterSetting { get; }

        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstColumn => FilterSetting.FirstColumn;

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int? LastColumn => FilterSetting.LastColumn;

        public FilterAttribute(int firstColumn) : this(firstColumn, null)
        {
        }

        public FilterAttribute(int firstColumn, int? lastColumn)
        {
            FilterSetting = new FilterSetting(firstColumn, lastColumn);
        }
    }
}
