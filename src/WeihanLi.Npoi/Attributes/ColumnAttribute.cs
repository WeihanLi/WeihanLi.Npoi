using System;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ColumnAttribute : Attribute
    {
        internal PropertySetting PropertySetting { get; }

        /// <summary>
        /// ColumnIndex
        /// </summary>
        public int Index
        {
            get => PropertySetting.ColumnIndex;
            set
            {
                if (value >= 0)
                {
                    PropertySetting.ColumnIndex = value;
                }
            }
        }

        /// <summary>
        /// ColumnTitle
        /// </summary>
        public string Title { get => PropertySetting.ColumnTitle; set => PropertySetting.ColumnTitle = value; }

        /// <summary>
        /// Formatter
        /// </summary>
        public string Formatter { get => PropertySetting.ColumnFormatter; set => PropertySetting.ColumnFormatter = value; }

        /// <summary>
        /// IsIgnored
        /// </summary>
        public bool IsIgnored
        {
            get => PropertySetting.IsIgnored;
            set => PropertySetting.IsIgnored = value;
        }

        public ColumnAttribute() => PropertySetting = new PropertySetting();

        public ColumnAttribute(string title) => PropertySetting = new PropertySetting
        {
            ColumnTitle = title
        };
    }
}
