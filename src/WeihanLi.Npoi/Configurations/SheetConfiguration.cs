using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class SheetConfiguration : ISheetConfiguration
    {
        internal SheetSetting SheetSetting { get; }

        public SheetConfiguration()
            => SheetSetting = new SheetSetting();

        public SheetConfiguration(SheetSetting sheetSetting)
            => SheetSetting = sheetSetting ?? new SheetSetting();

        public ISheetConfiguration HasSheetIndex(int index)
        {
            SheetSetting.SheetIndex = index;
            return this;
        }

        public ISheetConfiguration HasSheetName(string sheetName)
        {
            SheetSetting.SheetName = sheetName;
            return this;
        }

        public ISheetConfiguration HasStartRowIndex(int index)
        {
            SheetSetting.StartRowIndex = index;
            return this;
        }
    }
}
