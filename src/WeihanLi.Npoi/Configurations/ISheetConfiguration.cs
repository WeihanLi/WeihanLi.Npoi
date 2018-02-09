namespace WeihanLi.Npoi.Configurations
{
    /// <summary>
    /// SheetConfiguration
    /// </summary>
    public interface ISheetConfiguration
    {
        ISheetConfiguration HasSheetIndex(int index);

        ISheetConfiguration HasSheetName(string sheetName);

        ISheetConfiguration HasStartRowIndex(int index);
    }
}
