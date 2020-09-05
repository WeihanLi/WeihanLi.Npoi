namespace WeihanLi.Document.Abstract.Excel
{
    public interface IWorkbook
    {
        int SheetCount { get; }

        ISheet GetSheet(int sheetIndex);

        ISheet CreateSheet(string sheetName);

        byte[] ToBytes();
    }
}
