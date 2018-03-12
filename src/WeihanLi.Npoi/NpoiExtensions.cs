using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    public static class NpoiExtensions
    {
        /// <summary>
        /// Workbook2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="workbook">excel workbook</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]this IWorkbook workbook) where TEntity : new() => workbook.ToEntityList<TEntity>(0);

        /// <summary>
        /// Workbook2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]this IWorkbook workbook, int sheetIndex) where TEntity : new()
        {
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            return new NpoiHelper<TEntity>().SheetToEntityList(sheet);
        }

        /// <summary>
        /// Sheet2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">excel sheet</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]this ISheet sheet) where TEntity : new() => new NpoiHelper<TEntity>().SheetToEntityList(sheet);

        /// <summary>
        /// Workbook2ToDataTable
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]this IWorkbook workbook) => workbook.ToDataTable(0, 0);

        public static DataSet ToDataSet([NotNull] this IWorkbook workbook) => workbook.ToDataSet(0);

        public static DataSet ToDataSet([NotNull] this IWorkbook workbook, int headerRowIndex)
        {
            var ds = new DataSet();
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                ds.Tables.Add(workbook.GetSheetAt(i).ToDataTable(headerRowIndex));
            }
            return ds;
        }

        /// <summary>
        /// Workbook2ToDataTable
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]this IWorkbook workbook, int sheetIndex, int headerRowIndex)
        {
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            return workbook.GetSheetAt(sheetIndex).ToDataTable(headerRowIndex);
        }

        /// <summary>
        /// Sheet2DataTable
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]this ISheet sheet) => sheet.ToDataTable(0);

        /// <summary>
        /// Sheet2DataTable
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]this ISheet sheet, int headerRowIndex)
        {
            if (sheet.PhysicalNumberOfRows <= headerRowIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRowIndex), string.Format(Resource.IndexOutOfRange, nameof(headerRowIndex), sheet.PhysicalNumberOfRows));
            }
            var dataTable = new DataTable(sheet.SheetName);
            var rowEnumerator = sheet.GetRowEnumerator();
            while (rowEnumerator.MoveNext())
            {
                if (!(rowEnumerator.Current is IRow row))
                {
                    continue;
                }

                if (row.RowNum < headerRowIndex)
                {
                    continue;
                }

                if (row.RowNum == headerRowIndex)
                {
                    foreach (var cell in row.Cells)
                    {
                        dataTable.Columns.Add(cell.StringCellValue.Trim());
                    }
                }
                else
                {
                    var dataRow = dataTable.NewRow();
                    for (var i = 0; i < row.Cells.Count; i++)
                    {
                        dataRow[i] = row.Cells[i].GetCellValue(typeof(string));
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }
            return dataTable;
        }

        /// <summary>
        /// import entityList to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="list">entityList</param>
        public static int ImportData<TEntity>([NotNull]this IWorkbook workbook, IEnumerable<TEntity> list) where TEntity : new() => workbook.ImportData(list, 0);

        /// <summary>
        /// import entityList to workbook sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="list">entityList</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static int ImportData<TEntity>([NotNull]this IWorkbook workbook, IEnumerable<TEntity> list, int sheetIndex) where TEntity : new()
        {
            if (sheetIndex >= ExcelConstants.MaxSheetNum)
            {
                throw new ArgumentException(string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), ExcelConstants.MaxSheetNum), nameof(sheetIndex));
            }
            InternalCache.TypeExcelConfigurationDictionary.TryGetValue(typeof(TEntity), out var configuration);
            while (workbook.NumberOfSheets <= sheetIndex)
            {
                if (configuration?.SheetConfigurations == null || configuration.SheetConfigurations.Count <= sheetIndex)
                {
                    workbook.CreateSheet();
                }
                else
                {
                    workbook.CreateSheet((configuration.SheetConfigurations[sheetIndex] as SheetConfiguration)?.SheetSetting.SheetName ?? $"Sheet{sheetIndex}");
                }
            }
            new NpoiHelper<TEntity>().EntityListToSheet(workbook.GetSheetAt(sheetIndex), list.ToArray());
            return 1;
        }

        /// <summary>
        /// import entityList to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="list">entityList</param>
        public static ISheet ImportData<TEntity>([NotNull]this ISheet sheet, IEnumerable<TEntity> list)
            where TEntity : new() => new NpoiHelper<TEntity>().EntityListToSheet(sheet, list.ToArray());

        /// <summary>
        /// import datatable to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="dataTable">dataTable</param>
        public static int ImportData<TEntity>([NotNull]this IWorkbook workbook, [NotNull]DataTable dataTable) where TEntity : new() => workbook.ImportData<TEntity>(dataTable, 0);

        /// <summary>
        /// import datatable to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="dataTable">dataTable</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static int ImportData<TEntity>([NotNull]this IWorkbook workbook, [NotNull]DataTable dataTable, int sheetIndex) where TEntity : new()
        {
            if (sheetIndex >= ExcelConstants.MaxSheetNum)
            {
                throw new ArgumentException(string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), ExcelConstants.MaxSheetNum), nameof(sheetIndex));
            }
            InternalCache.TypeExcelConfigurationDictionary.TryGetValue(typeof(TEntity), out var configuration);
            while (workbook.NumberOfSheets <= sheetIndex)
            {
                if (configuration?.SheetConfigurations == null || configuration.SheetConfigurations.Count <= sheetIndex)
                {
                    workbook.CreateSheet();
                }
                else
                {
                    workbook.CreateSheet((configuration.SheetConfigurations[sheetIndex] as SheetConfiguration)?.SheetSetting.SheetName ?? $"Sheet{sheetIndex}");
                }
            }
            new NpoiHelper<TEntity>().DataTableToSheet(workbook.GetSheetAt(sheetIndex), dataTable);
            return 1;
        }

        /// <summary>
        /// import datatable to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="dataTable">dataTable</param>
        public static ISheet ImportData<TEntity>([NotNull]this ISheet sheet, DataTable dataTable)
            where TEntity : new() => new NpoiHelper<TEntity>().DataTableToSheet(sheet, dataTable);

        /// <summary>
        /// EntityList2ExcelFile
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelPath">excelPath</param>
        public static int ToExcelFile<TEntity>([NotNull] this IEnumerable<TEntity> entityList, [NotNull]string excelPath)
            where TEntity : new()
        {
            InternalCache.TypeExcelConfigurationDictionary.TryGetValue(typeof(TEntity), out var configuration);
            var workbook = ExcelHelper.PrepareWorkbook(excelPath, configuration?.ExcelSetting);
            workbook.ImportData(entityList.ToArray());
            workbook.WriteToFile(excelPath);
            return 1;
        }

        /// <summary>
        /// EntityList2ExcelStream
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="stream">stream where to write</param>
        public static int ToExcelStream<TEntity>([NotNull] this IEnumerable<TEntity> entityList, [NotNull]Stream stream)
            where TEntity : new()
        {
            InternalCache.TypeExcelConfigurationDictionary.TryGetValue(typeof(TEntity), out var configuration);
            var workbook = ExcelHelper.PrepareWorkbook(true, configuration?.ExcelSetting);
            workbook.ImportData(entityList.ToArray());
            workbook.Write(stream);
            return 1;
        }

        /// <summary>
        /// EntityList2ExcelBytes
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        public static byte[] ToExcelBytes<TEntity>([NotNull] this IEnumerable<TEntity> entityList)
            where TEntity : new()
        {
            InternalCache.TypeExcelConfigurationDictionary.TryGetValue(typeof(TEntity), out var configuration);
            var workbook = ExcelHelper.PrepareWorkbook(true, configuration?.ExcelSetting);
            workbook.ImportData(entityList.ToArray());
            return workbook.ToExcelBytes();
        }

        /// <summary>
        /// 将DataTable内容导出到Excel
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="excelPath">excelPath</param>
        /// <returns></returns>
        public static int ToExcelFile([NotNull] this DataTable dataTable, [NotNull] string excelPath)
        {
            var workbook = ExcelHelper.PrepareWorkbook(excelPath);
            var sheet = workbook.CreateSheet(string.IsNullOrWhiteSpace(dataTable.TableName) ? "Sheet0" : dataTable.TableName);
            var headerRow = sheet.CreateRow(0);
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i, CellType.String).SetCellValue(dataTable.Columns[i].ColumnName);
            }

            for (var i = 1; i <= dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(dataTable.Rows[i - 1][j]);
                }
            }

            workbook.WriteToFile(excelPath);
            return 1;
        }

        /// <summary>
        /// DataTable2ExcelStream
        /// </summary>
        /// <param name="dataTable">datatable</param>
        /// <param name="stream">stream</param>
        /// <returns></returns>
        public static int ToExcelStream([NotNull] this DataTable dataTable, [NotNull] Stream stream)
        {
            var workbook = ExcelHelper.PrepareWorkbook();
            var sheet = workbook.CreateSheet(string.IsNullOrWhiteSpace(dataTable.TableName) ? "Sheet0" : dataTable.TableName);
            var headerRow = sheet.CreateRow(0);
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i, CellType.String).SetCellValue(dataTable.Columns[i].ColumnName);
            }

            for (var i = 1; i <= dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(dataTable.Rows[i][j]);
                }
            }

            workbook.Write(stream);
            return 1;
        }

        /// <summary>
        /// DataTable2ExcelBytes
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        public static byte[] ToExcelBytes([NotNull] this DataTable dataTable)
        {
            var workbook = ExcelHelper.PrepareWorkbook();
            var sheet = workbook.CreateSheet(string.IsNullOrWhiteSpace(dataTable.TableName) ? "Sheet0" : dataTable.TableName);
            var headerRow = sheet.CreateRow(0);
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i, CellType.String).SetCellValue(dataTable.Columns[i].ColumnName);
            }

            for (var i = 1; i <= dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(dataTable.Rows[i][j]);
                }
            }

            return workbook.ToExcelBytes();
        }

        /// <summary>
        /// SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        public static void SetCellValue([NotNull]this ICell cell, object value) => cell.SetCellValue(value, null);

        /// <summary>
        /// SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        /// <param name="formatter">formatter</param>
        public static void SetCellValue([NotNull]this ICell cell, object value, string formatter)
        {
            if (null == value)
            {
                cell.SetCellType(CellType.Blank);
                return;
            }
            if (value is DateTime time)
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(string.IsNullOrWhiteSpace(formatter) ?
                    (time.Date == time ? time.ToStandardDateString() : time.ToStandardTimeString())
                    : time.ToString(formatter));
            }
            else
            {
                var type = value.GetType();
                if (
                    type == typeof(double) ||
                    type == typeof(int) ||
                    type == typeof(long) ||
                    type == typeof(float) ||
                    type == typeof(decimal)
                )
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(Convert.ToDouble(value));
                }
                else if (type == typeof(bool))
                {
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue((bool)value);
                }
                else
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(value is IFormattable val && formatter.IsNotNullOrWhiteSpace() ? val.ToString(formatter, CultureInfo.CurrentCulture) : value.ToString());
                }
            }

            //try
            //{
            //    if (formatter.IsNotNullOrWhiteSpace())
            //    {
            //        cell.CellStyle.DataFormat = cell.Row.Sheet.Workbook.CreateDataFormat().GetFormat(formatter);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Diagnostics.Debug.WriteLine(ex.ToString());
            //}
        }

        /// <summary>
        /// GetCellValue
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="propertyType">propertyType</param>
        /// <returns>cellValue</returns>
        public static object GetCellValue([NotNull]this ICell cell, Type propertyType)
        {
            if (string.IsNullOrEmpty(cell.ToString()))
            {
                return propertyType.GetDefaultValue();
            }
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        if (propertyType == typeof(DateTime))
                        {
                            return cell.DateCellValue;
                        }
                        return DateTime.Parse(cell.DateCellValue == cell.DateCellValue.Date
                            ? cell.DateCellValue.ToStandardDateString()
                            : cell.DateCellValue.ToStandardTimeString());
                    }
                    if (propertyType == typeof(double))
                    {
                        return cell.NumericCellValue;
                    }
                    // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                    // System.ComponentModel.TypeDescriptor.GetConverter(typeof(decimal)).CanConvertFrom(typeof(double)) :: false
                    // HACK:直接 ToOrDefault() 时，double 转换为 decimal 转换后还是 double
                    return cell.NumericCellValue.ToString().ToOrDefault(propertyType);

                case CellType.String:
                    if (propertyType == typeof(string))
                    {
                        return cell.StringCellValue;
                    }
                    return cell.StringCellValue.ToOrDefault(propertyType);

                case CellType.Boolean:
                    if (propertyType == typeof(bool))
                    {
                        return cell.BooleanCellValue;
                    }
                    return cell.BooleanCellValue.ToString().ToOrDefault(propertyType);

                case CellType.Blank:
                    return propertyType.GetDefaultValue();

                default:
                    return cell.ToString().ToOrDefault(propertyType);
            }
        }

        /// <summary>
        /// GetCellValue
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="cell">cell</param>
        /// <returns></returns>
        public static T GetCellValue<T>([NotNull] this ICell cell)
        {
            if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            {
                return cell.DateCellValue.Date == cell.DateCellValue ? cell.DateCellValue.ToStandardDateString().ToOrDefault<T>() : cell.DateCellValue.ToStandardTimeString().ToOrDefault<T>();
            }
            return cell.ToString().ToOrDefault<T>();
        }

        /// <summary>
        /// Write workbook to excel file
        /// </summary>
        /// <param name="workbook">workbook</param>
        /// <param name="filePath">file path</param>
        public static void WriteToFile([NotNull] this IWorkbook workbook, string filePath)
        {
            using (var fileStream = File.Create(filePath))
            {
                workbook.Write(fileStream);
            }
        }

        /// <summary>
        /// ToExcelBytes
        /// </summary>
        /// <param name="workbook">workbook</param>
        public static byte[] ToExcelBytes([NotNull] this IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }
    }
}
