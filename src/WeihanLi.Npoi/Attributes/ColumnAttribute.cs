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

        public ColumnAttribute(string title) => Title = title;
    }
}
