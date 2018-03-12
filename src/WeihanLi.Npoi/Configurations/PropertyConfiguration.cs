using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        internal PropertyConfiguration() => PropertySetting = new PropertySetting();

        internal PropertyConfiguration(PropertySetting propertySetting) => PropertySetting = propertySetting;

        internal PropertySetting PropertySetting { get; }

        public IPropertyConfiguration HasColumnIndex(int index)
        {
            if (index < 0)
            {
                PropertySetting.ColumnIndex = index;
            }
            return this;
        }

        public IPropertyConfiguration HasColumnTitle(string title)
        {
            if (title.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnTitle = title;
            }
            return this;
        }

        public IPropertyConfiguration HasColumnFormatter(string formatter)
        {
            if (formatter.IsNotNullOrWhiteSpace())
            {
                PropertySetting.ColumnFormatter = formatter;
            }
            return this;
        }

        public IPropertyConfiguration Ignored()
        {
            PropertySetting.IsIgnored = true;
            return this;
        }
    }
}
