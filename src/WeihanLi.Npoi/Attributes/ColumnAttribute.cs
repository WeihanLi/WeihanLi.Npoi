using System;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ColumnAttribute : Attribute
    {
        /// <summary>
        /// ColumnIndex
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// ColumnTitle
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Formatter
        /// </summary>
        public string Formatter { get; set; }

        public ColumnAttribute()
        {
        }

        public ColumnAttribute(string title) => Title = title;
    }
}
