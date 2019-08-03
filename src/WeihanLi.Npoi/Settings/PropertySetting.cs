using System;

namespace WeihanLi.Npoi.Settings
{
    /// <summary>
    /// PropertySetting
    /// </summary>
    internal class PropertySetting
    {
        /// <summary>
        /// ColumnIndex
        /// </summary>
        public int ColumnIndex { get; set; } = -1;

        /// <summary>
        /// Title
        /// </summary>
        public string ColumnTitle { get; set; }

        /// <summary>
        /// Formatter
        /// </summary>
        public string ColumnFormatter { get; set; }

        /// <summary>
        /// the property is ignored.
        /// </summary>
        public bool IsIgnored { get; set; }
    }

    internal class PropertySetting<TEntity, TProperty> : PropertySetting
    {
        /// <summary>
        /// ColumnFormatterFunc
        /// </summary>
        public Func<TEntity, TProperty, object> ColumnFormatterFunc { get; set; }
    }
}
