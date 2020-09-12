using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using WeihanLi.Document.Abstract.Excel;

namespace WeihanLi.Epplus
{
    internal class EPPlusWorkbook : IWorkbook
    {
        private readonly ExcelPackage _excelPackage;

        public EPPlusWorkbook(ExcelPackage excelPackage)
        {
            _excelPackage = excelPackage;
        }

        public int SheetCount => _excelPackage.Workbook.Worksheets.Count;

        public ISheet CreateSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public ISheet GetSheet(int sheetIndex)
        {
            return new EPPlusSheet(_excelPackage.Workbook.Worksheets[sheetIndex]);
        }

        public byte[] ToBytes()
        {
            return _excelPackage.GetAsByteArray();
        }
    }

    internal class EPPlusSheet : ISheet
    {
        private readonly ExcelWorksheet _worksheet;

        public EPPlusSheet(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public int FirstRowNum { get; }
        public int LastRowNum { get; }

        public IRow GetRow(int rowIndex) => throw new System.NotImplementedException();

        public IRow CreateRow(int rowIndex) => throw new System.NotImplementedException();

        public void SetColumnWidth(int columnIndex, int width) => throw new System.NotImplementedException();

        public void AutoSizeColumn(int columnIndex) => throw new System.NotImplementedException();

        public void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow) => throw new System.NotImplementedException();

        public void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) => throw new System.NotImplementedException();

        public void ShiftRows(int startRow, int endRow, int n) => throw new System.NotImplementedException();

        public IRow CopyRow(int sourceIndex, int targetIndex) => throw new System.NotImplementedException();

        public void RemoveRow(IRow row) => throw new System.NotImplementedException();
    }

    internal class EPPlusRow : IRow
    {
        private readonly ExcelRow _row;

        public EPPlusRow(ExcelRow row)
        {
            _row = row;
        }

        public int CellsCount { get; }
        public int FirstCellNum { get; }
        public int LastCellNum { get; }

        public ICell GetCell(int cellIndex) => throw new System.NotImplementedException();

        public ICell CreateCell(int cellIndex) => throw new System.NotImplementedException();

        public object UnderlyingValue { get; }
    }

    internal class EPPlusCell : ICell
    {
        private readonly ExcelCell _cell;

        public EPPlusCell(ExcelCell cell)
        {
            _cell = cell;
        }

        public CellType CellType { get; set; }

        public object Value
        {
            get => _cell.Value;
            set { }
        }
    }
}
