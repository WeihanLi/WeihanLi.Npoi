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

    internal class PropertyConfiguration<TProperty> : PropertyConfiguration, IPropertyConfiguration<TProperty>
    {
        internal PropertyConfiguration() : this(new PropertySetting<TProperty>())
        {
        }

        public PropertyConfiguration(PropertySetting<TProperty> propertySetting) : base(propertySetting) => PropertySetting = propertySetting;

        internal new PropertySetting<TProperty> PropertySetting { get; }

        public IPropertyConfiguration<TProperty> HasColumnIndex(int index)
        {
            if (index >= 0)
            {
                PropertySetting.ColumnIndex = index;
            }
            return this;
        }

        public IPropertyConfiguration<TProperty> HasColumnTitle(string title)
        {
            if (title.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnTitle = title;
            }
            return this;
        }

        public IPropertyConfiguration<TProperty> HasColumnFormatter(string formatter)
        {
            if (formatter.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnFormatter = formatter;
            }
            return this;
        }

        public IPropertyConfiguration<TProperty> Ignored()
        {
            PropertySetting.IsIgnored = true;
            return this;
        }

        public IPropertyConfiguration<TProperty> HasColumnFormatter(Func<TProperty, object> formatterFunc)
        {
            if (formatterFunc != null)
            {
                PropertySetting.ColumnFormatterFunc = formatterFunc;
            }
            return this;
        }
    }
}
