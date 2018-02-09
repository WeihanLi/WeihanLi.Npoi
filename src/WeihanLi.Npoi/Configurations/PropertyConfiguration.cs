using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        internal PropertyConfiguration()
        {
            PropertySetting = new PropertySetting();
        }

        internal PropertySetting PropertySetting { get; }

        public IPropertyConfiguration HasColumnIndex(int index)
        {
            PropertySetting.ColumnIndex = index;
            return this;
        }

        public IPropertyConfiguration HasColumnTitle(string title)
        {
            PropertySetting.ColumnTitle = title;
            return this;
        }

        public IPropertyConfiguration HasColumnFormatter(string formatter)
        {
            PropertySetting.ColumnFormatter = formatter;
            return this;
        }
    }
}
