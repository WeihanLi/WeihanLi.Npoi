using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class SheetConfiguration : ISheetConfiguration
    {
        internal SheetSetting SheetSetting { get; }

        public SheetConfiguration() : this(null)
        {
        }

        public SheetConfiguration(SheetSetting sheetSetting)
            => SheetSetting = sheetSetting ?? new SheetSetting();

        public ISheetConfiguration HasSheetIndex(int index)
        {
            if (index >= 0)
            {
                SheetSetting.SheetIndex = index;
            }
            return this;
        }

        public ISheetConfiguration HasSheetName(string sheetName)
        {
            if (sheetName.IsNotNullOrWhiteSpace())
            {
                SheetSetting.SheetName = sheetName;
            }
            return this;
        }

        public ISheetConfiguration HasStartRowIndex(int index)
        {
            SheetSetting.StartRowIndex = index;
            return this;
        }
    }
}
