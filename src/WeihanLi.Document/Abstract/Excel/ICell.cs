namespace WeihanLi.Document.Abstract.Excel
{
    public interface ICell
    {
        CellType CellType { get; set; }

        object Value { get; set; }
    }

    public enum CellType
    {
        Unknown = -1, // 0xFFFFFFFF
        Numeric = 0,
        String = 1,
        Formula = 2,
        Blank = 3,
        Boolean = 4,
        Error = 5,
    }
}
