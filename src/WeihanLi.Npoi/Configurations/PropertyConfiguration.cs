using System;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        internal PropertySetting PropertySetting { get; }

        public PropertyConfiguration(PropertySetting propertySetting) => PropertySetting = propertySetting;
    }

    internal class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration, IPropertyConfiguration<TEntity, TProperty>
    {
        internal PropertyConfiguration() : this(new PropertySetting<TEntity, TProperty>())
        {
        }

        public PropertyConfiguration(PropertySetting<TEntity, TProperty> propertySetting) : base(propertySetting) => PropertySetting = propertySetting;

        internal new PropertySetting<TEntity, TProperty> PropertySetting { get; }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index)
        {
            if (index >= 0)
            {
                PropertySetting.ColumnIndex = index;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title)
        {
            if (title.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnTitle = title;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width)
        {
            PropertySetting.ColumnWidth = width;
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter)
        {
            if (formatter.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnFormatter = formatter;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> Ignored()
        {
            PropertySetting.IsIgnored = true;
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            if (formatterFunc != null)
            {
                PropertySetting.ColumnFormatterFunc = formatterFunc;
            }
            return this;
        }
    }
}
