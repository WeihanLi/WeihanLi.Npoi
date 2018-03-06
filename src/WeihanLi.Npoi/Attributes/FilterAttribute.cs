using System;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class FilterAttribute : Attribute
    {
        internal FilterSetting FilterSeting { get; }

        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstColumn => FilterSeting.FirstColumn;

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int? LastColumn => FilterSeting.LastColumn;

        public FilterAttribute(int firstColumn) : this(firstColumn, null)
        {
        }

        public FilterAttribute(int firstColumn, int? lastColumn)
        {
            FilterSeting = new FilterSetting(firstColumn, lastColumn);
        }
    }
}
