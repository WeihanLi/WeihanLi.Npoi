using NPOI.SS.Util;
using System;
using NModel = NPOI.SS.UserModel;

namespace WeihanLi.Npoi.Abstract
{
    internal class NPOIWorkbook : IWorkbook
    {
        public int SheetCount => _workbook.NumberOfSheets;

        public ISheet GetSheet(int sheetIndex) => new NPOISheet(_workbook.GetSheetAt(sheetIndex));

        public ISheet CreateSheet(string sheetName) => new NPOISheet(_workbook.CreateSheet(sheetName));

        public byte[] ToBytes() => _workbook.ToExcelBytes();

        private readonly NModel.IWorkbook _workbook;

        public NPOIWorkbook(NModel.IWorkbook workbook)
        {
            _workbook = workbook;
        }
    }

    internal class NPOISheet : ISheet
    {
        public int FirstRowNum => _sheet.FirstRowNum + 1;
        public int LastRowNum => _sheet.LastRowNum + 1;

        public IRow GetRow(int rowIndex)
        {
            var nRow = _sheet.GetRow(rowIndex);
            if (null == nRow)
                return null;

            return new NPOIRow(nRow);
        }

        public IRow CreateRow(int rowIndex) => new NPOIRow(_sheet.CreateRow(rowIndex));

        public void SetColumnWidth(int columnIndex, int width) => _sheet.SetColumnWidth(columnIndex, width);

        public void AutoSizeColumn(int columnIndex) => _sheet.AutoSizeColumn(columnIndex);

        public void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow) => _sheet.CreateFreezePane(colSplit, rowSplit, leftMostCol, topRow);

        public void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) => _sheet.SetAutoFilter(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex));

        public void ShiftRows(int startRow, int endRow, int n) => _sheet.ShiftRows(startRow, endRow, n);

        public IRow CopyRow(int sourceIndex, int targetIndex) => new NPOIRow(_sheet.CopyRow(sourceIndex, targetIndex));

        public void RemoveRow(IRow row) => _sheet.RemoveRow(row.UnderlyingValue as NModel.IRow);

        private readonly NModel.ISheet _sheet;

        public NPOISheet(NModel.ISheet sheet)
        {
            _sheet = sheet;
        }
    }

    internal class NPOIRow : IRow
    {
        public int CellsCount => _row.PhysicalNumberOfCells;
        public int FirstCellNum => _row.FirstCellNum + 1;
        public int LastCellNum => _row.LastCellNum;

        public ICell GetCell(int cellIndex)
        {
            var nCell = _row.GetCell(cellIndex);
            if (nCell is null) return null;
            return new NPOICell(nCell);
        }

        public ICell CreateCell(int cellIndex) => new NPOICell(_row.CreateCell(cellIndex));

        public object UnderlyingValue => _row;

        private readonly NModel.IRow _row;

        public NPOIRow(NModel.IRow row)
        {
            _row = row;
        }
    }

    internal class NPOICell : ICell
    {
        private readonly NModel.ICell _cell;

        public NPOICell(NModel.ICell cell)
        {
            _cell = cell;
        }

        public CellType CellType
        {
            get => (CellType)Enum.Parse(typeof(CellType), _cell.CellType.ToString());
            set => _cell.SetCellType((NModel.CellType)Enum.Parse(typeof(NModel.CellType), value.ToString()));
        }

        public object Value
        {
            get
            {
                if (_cell is null || _cell.CellType == NModel.CellType.Blank || _cell.CellType == NModel.CellType.Error)
                {
                    return null;
                }
                switch (_cell.CellType)
                {
                    case NModel.CellType.Numeric:
                        if (NModel.DateUtil.IsCellDateFormatted(_cell))
                        {
                            return _cell.DateCellValue;
                        }
                        return _cell.NumericCellValue;

                    case NModel.CellType.String:
                        return _cell.StringCellValue;

                    case NModel.CellType.Boolean:
                        return _cell.BooleanCellValue;

                    default:
                        return _cell.ToString();
                }
            }

            set => _cell.SetCellValue(value);
        }
    }
}
