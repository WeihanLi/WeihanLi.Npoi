using JetBrains.Annotations;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    public static class NpoiExtensions
    {
        /// <summary>
        ///     Workbook2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="workbook">excel workbook</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] this IWorkbook workbook) where TEntity : new() => workbook.ToEntityList<TEntity>(0);

        /// <summary>
        ///     Workbook2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] this IWorkbook workbook, int sheetIndex)
            where TEntity : new()
        {
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex),
                    string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            return NpoiHelper.SheetToEntityList<TEntity>(sheet, sheetIndex);
        }

        /// <summary>
        ///     Sheet2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">excel sheet</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] this ISheet sheet) where TEntity : new() => sheet.ToEntityList<TEntity>(0);

        /// <summary>
        ///     Sheet2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">excel sheet</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] this ISheet sheet, int sheetIndex)
            where TEntity : new() => NpoiHelper.SheetToEntityList<TEntity>(sheet, sheetIndex);

        /// <summary>
        ///     Workbook2ToDataTable
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull] this IWorkbook workbook) => workbook.ToDataTable(0, 0);

        /// <summary>
        ///     Workbook2ToDataSet
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <returns>DataSet</returns>
        public static DataSet ToDataSet([NotNull] this IWorkbook workbook) => workbook.ToDataSet(0);

        /// <summary>
        ///     Workbook2ToDataSet
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataSet</returns>
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
        ///     Workbook2ToDataTable
        /// </summary>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull] this IWorkbook workbook, int sheetIndex, int headerRowIndex)
        {
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex),
                    string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            return workbook.GetSheetAt(sheetIndex).ToDataTable(headerRowIndex);
        }

        /// <summary>
        ///     Sheet2DataTable
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull] this ISheet sheet) => sheet.ToDataTable(0);

        /// <summary>
        ///     Sheet2DataTable
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull] this ISheet sheet, int headerRowIndex)
        {
            if (sheet.LastRowNum <= headerRowIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRowIndex),
                    string.Format(Resource.IndexOutOfRange, nameof(headerRowIndex), sheet.PhysicalNumberOfRows));
            }
            var dataTable = new DataTable(sheet.SheetName);

            foreach (var row in sheet.GetRowCollection())
            {
                if (row == null || row.RowNum < headerRowIndex)
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
        ///     import entityList to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="list">entityList</param>
        public static int ImportData<TEntity>([NotNull] this IWorkbook workbook, IEnumerable<TEntity> list)
             => workbook.ImportData(list, 0);

        /// <summary>
        ///     import entityList to workbook sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="list">entityList</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>the sheet LastRowNum</returns>
        public static int ImportData<TEntity>([NotNull] this IWorkbook workbook, IEnumerable<TEntity> list,
            int sheetIndex)
        {
            if (sheetIndex >= InternalConstants.MaxSheetNum)
            {
                throw new ArgumentException(
                    string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), InternalConstants.MaxSheetNum),
                    nameof(sheetIndex));
            }

            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();

            while (workbook.NumberOfSheets <= sheetIndex)
            {
                if (workbook.NumberOfSheets == sheetIndex)
                {
                    var sheetName = typeof(TEntity).Name;
                    if (configuration.SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
                    {
                        sheetName = sheetSetting.SheetName;
                    }
                    workbook.CreateSheet(sheetName);
                }
                else
                {
                    workbook.CreateSheet();
                }
            }

            var sheet = NpoiHelper.EntityListToSheet(workbook.GetSheetAt(sheetIndex), list, sheetIndex);
            return sheet.LastRowNum;
        }

        /// <summary>
        ///     import entityList to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="list">entityList</param>
        public static ISheet ImportData<TEntity>([NotNull] this ISheet sheet, IEnumerable<TEntity> list)
             => sheet.ImportData(list, 0);

        /// <summary>
        ///     import entityList to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="list">entityList</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static ISheet ImportData<TEntity>([NotNull] this ISheet sheet, IEnumerable<TEntity> list, int sheetIndex)
             => NpoiHelper.EntityListToSheet(sheet, list, sheetIndex);

        /// <summary>
        ///     import dataTable to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="dataTable">dataTable</param>
        public static int ImportData<TEntity>([NotNull] this IWorkbook workbook, [NotNull] DataTable dataTable)
             => workbook.ImportData<TEntity>(dataTable, 0);

        /// <summary>
        ///     import dataTable to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="dataTable">dataTable</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>the sheet LastRowNum</returns>
        public static int ImportData<TEntity>([NotNull] this IWorkbook workbook, [NotNull] DataTable dataTable,
            int sheetIndex)
        {
            if (sheetIndex >= InternalConstants.MaxSheetNum)
            {
                throw new ArgumentException(
                    string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), InternalConstants.MaxSheetNum),
                    nameof(sheetIndex));
            }

            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();

            while (workbook.NumberOfSheets <= sheetIndex)
            {
                if (workbook.NumberOfSheets == sheetIndex)
                {
                    var sheetName = typeof(TEntity).Name;
                    if (configuration.SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
                    {
                        sheetName = sheetSetting.SheetName;
                    }
                    workbook.CreateSheet(sheetName);
                }
                else
                {
                    workbook.CreateSheet();
                }
            }

            var sheet = NpoiHelper.DataTableToSheet<TEntity>(workbook.GetSheetAt(sheetIndex), dataTable, sheetIndex);
            return sheet.LastRowNum;
        }

        /// <summary>
        /// import dataTable to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="dataTable">dataTable</param>
        public static ISheet ImportData<TEntity>([NotNull] this ISheet sheet, DataTable dataTable) => sheet.ImportData<TEntity>(dataTable, 0);

        /// <summary>
        /// import dataTable to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="dataTable">dataTable</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static ISheet ImportData<TEntity>([NotNull] this ISheet sheet, DataTable dataTable, int sheetIndex)
             => NpoiHelper.DataTableToSheet<TEntity>(sheet, dataTable, sheetIndex);

        #region ExportByTemplate

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templatePath"></param>
        /// <param name="excelPath">templateBytes</param>
        /// <param name="sheetIndex">sheetIndex,zero by default</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static void ToExcelFileByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, string templatePath, string excelPath, int sheetIndex = 0, object extraData = null)
        {
            if (templatePath == null)
            {
                throw new ArgumentNullException(nameof(templatePath));
            }
            if (excelPath == null)
            {
                throw new ArgumentNullException(nameof(excelPath));
            }

            var workbook = ExcelHelper.LoadExcel(templatePath);
            entities.ToExcelFileByTemplate(workbook, excelPath, sheetIndex, extraData);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateBytes">templateBytes</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex,zero by default</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static void ToExcelFileByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, byte[] templateBytes, string excelPath, ExcelFormat excelFormat = ExcelFormat.Xls, int sheetIndex = 0, object extraData = null)
        {
            if (templateBytes == null)
            {
                throw new ArgumentNullException(nameof(templateBytes));
            }
            if (excelPath == null)
            {
                throw new ArgumentNullException(nameof(excelPath));
            }

            var workbook = ExcelHelper.LoadExcel(templateBytes, excelFormat);
            entities.ToExcelFileByTemplate(workbook, excelPath, sheetIndex, extraData);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateWorkbook">templateWorkbook</param>
        /// <param name="excelPath"></param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static void ToExcelFileByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, IWorkbook templateWorkbook, string excelPath, int sheetIndex = 0, object extraData = null)
        {
            if (null == templateWorkbook)
            {
                throw new ArgumentNullException(nameof(templateWorkbook));
            }

            if (sheetIndex <= 0)
            {
                sheetIndex = 0;
            }

            var templateSheet = templateWorkbook.GetSheetAt(sheetIndex);
            NpoiTemplateHelper.EntityListToSheetByTemplate(
                templateSheet, entities, extraData
            );
            templateWorkbook.WriteToFile(excelPath);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templatePath">templatePath</param>
        /// <param name="sheetIndex">sheetIndex,zero by default</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static byte[] ToExcelBytesByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, string templatePath, int sheetIndex = 0, object extraData = null)
        {
            return ToExcelBytesByTemplate(entities, ExcelHelper.LoadExcel(templatePath), sheetIndex, extraData);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateBytes">templateBytes</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex,zero by default</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static byte[] ToExcelBytesByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, byte[] templateBytes, ExcelFormat excelFormat = ExcelFormat.Xls, int sheetIndex = 0, object extraData = null)
        {
            if (templateBytes == null)
            {
                throw new ArgumentNullException(nameof(templateBytes));
            }

            var workbook = ExcelHelper.LoadExcel(templateBytes, excelFormat);
            return ToExcelBytesByTemplate(entities, workbook, sheetIndex, extraData);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateStream">templateStream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex,zero by default</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static byte[] ToExcelBytesByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, Stream templateStream, ExcelFormat excelFormat = ExcelFormat.Xls, int sheetIndex = 0, object extraData = null)
        {
            if (templateStream == null)
            {
                throw new ArgumentNullException(nameof(templateStream));
            }

            var workbook = ExcelHelper.LoadExcel(templateStream, excelFormat);
            return ToExcelBytesByTemplate(entities, workbook, sheetIndex, extraData);
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateWorkbook">templateWorkbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static byte[] ToExcelBytesByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, IWorkbook templateWorkbook, int sheetIndex = 0, object extraData = null)
        {
            if (null == templateWorkbook)
            {
                throw new ArgumentNullException(nameof(templateWorkbook));
            }

            if (sheetIndex <= 0)
            {
                sheetIndex = 0;
            }
            var templateSheet = templateWorkbook.GetSheetAt(sheetIndex);
            NpoiTemplateHelper.EntityListToSheetByTemplate(
                templateSheet, entities, extraData
            );
            return templateWorkbook.ToExcelBytes();
        }

        /// <summary>
        /// export excel via template
        /// </summary>
        /// <typeparam name="TEntity">Entity Type</typeparam>
        /// <param name="entities">entities</param>
        /// <param name="templateSheet"></param>
        /// <param name="extraData">extraData</param>
        /// <returns>exported excel bytes</returns>
        public static byte[] ToExcelBytesByTemplate<TEntity>([NotNull]this IEnumerable<TEntity> entities, ISheet templateSheet, object extraData = null)
        {
            NpoiTemplateHelper.EntityListToSheetByTemplate(
                templateSheet, entities, extraData
            );
            return templateSheet.Workbook.ToExcelBytes();
        }

        #endregion ExportByTemplate

        /// <summary>
        ///     EntityList2ExcelFile
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelPath">excelPath</param>
        public static void ToExcelFile<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] string excelPath) => ToExcelFile(entityList, excelPath, 0);

        /// <summary>
        ///     EntityList2ExcelFile
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static void ToExcelFile<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] string excelPath, int sheetIndex)

        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();

            var workbook = ExcelHelper.PrepareWorkbook(excelPath, configuration.ExcelSetting);
            workbook.ImportData(entityList.ToArray(), sheetIndex);

            workbook.WriteToFile(excelPath);
        }

        /// <summary>
        ///     EntityList2ExcelStream(*.xls by default)
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="stream">stream where to write</param>
        public static void ToExcelStream<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] Stream stream) => ToExcelStream(entityList, stream, ExcelFormat.Xls);

        /// <summary>
        ///     EntityList2ExcelStream
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="stream">stream where to write</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static void ToExcelStream<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] Stream stream, ExcelFormat excelFormat, int sheetIndex)
        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();

            var workbook = ExcelHelper.PrepareWorkbook(excelFormat, configuration.ExcelSetting);
            workbook.ImportData(entityList.ToArray(), sheetIndex);
            workbook.Write(stream);
        }

        /// <summary>
        ///     EntityList2ExcelStream
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="stream">stream where to write</param>
        /// <param name="excelFormat">excelFormat</param>
        public static void ToExcelStream<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] Stream stream, ExcelFormat excelFormat) => ToExcelStream(entityList, stream, excelFormat, 0);

        /// <summary>
        ///     EntityList2ExcelBytes(*.xls by default)
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        public static byte[] ToExcelBytes<TEntity>([NotNull] this IEnumerable<TEntity> entityList) =>
            ToExcelBytes(entityList, ExcelFormat.Xls);

        /// <summary>
        ///     EntityList2ExcelBytes
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelFormat">excelFormat</param>
        public static byte[] ToExcelBytes<TEntity>([NotNull] this IEnumerable<TEntity> entityList, ExcelFormat excelFormat)
            => ToExcelBytes(entityList, excelFormat, 0);

        /// <summary>
        ///     EntityList2ExcelBytes
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static byte[] ToExcelBytes<TEntity>([NotNull] this IEnumerable<TEntity> entityList, ExcelFormat excelFormat, int sheetIndex)

        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();

            var workbook = ExcelHelper.PrepareWorkbook(excelFormat, configuration.ExcelSetting);
            workbook.ImportData(entityList.ToArray(), sheetIndex);

            return workbook.ToExcelBytes();
        }

        /// <summary>
        ///     export DataTable to excel file
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="excelPath">excelPath</param>
        /// <returns></returns>
        public static void ToExcelFile([NotNull] this DataTable dataTable, [NotNull] string excelPath) => ToExcelFile(dataTable, excelPath, null);

        /// <summary>
        ///     export DataTable to excel file
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="excelPath">excelPath</param>
        /// <param name="excelSetting">excelSetting</param>
        /// <returns></returns>
        public static void ToExcelFile([NotNull] this DataTable dataTable, [NotNull] string excelPath, ExcelSetting excelSetting)
        {
            var workbook = ExcelHelper.PrepareWorkbook(excelPath, excelSetting);
            var sheet = workbook.CreateSheet(
                string.IsNullOrWhiteSpace(dataTable.TableName)
                ? "Sheet0"
                : dataTable.TableName);
            if (dataTable.Columns.Count > 0)
            {
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
            }
            workbook.WriteToFile(excelPath);
        }

        /// <summary>
        ///     DataTable2ExcelStream
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="stream">stream</param>
        /// <returns></returns>
        public static void ToExcelStream([NotNull] this DataTable dataTable, [NotNull] Stream stream) => ToExcelStream(dataTable, stream, ExcelFormat.Xls);

        /// <summary>
        ///     DataTable2ExcelStream
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="stream">stream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns></returns>
        public static void ToExcelStream([NotNull] this DataTable dataTable, [NotNull] Stream stream, ExcelFormat excelFormat) => ToExcelStream(dataTable, stream, excelFormat, null);

        /// <summary>
        ///     DataTable2ExcelStream
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="stream">stream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="excelSetting">excelSetting</param>
        /// <returns></returns>
        public static void ToExcelStream([NotNull] this DataTable dataTable, [NotNull] Stream stream, ExcelFormat excelFormat, ExcelSetting excelSetting)

        {
            var workbook = ExcelHelper.PrepareWorkbook(excelFormat);
            var sheet = workbook.CreateSheet(
                string.IsNullOrWhiteSpace(dataTable.TableName)
                ? "Sheet0"
                : dataTable.TableName);
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

            workbook.Write(stream);
        }

        /// <summary>
        ///     DataTable2ExcelBytes(*.xlsx by default)
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        public static byte[] ToExcelBytes([NotNull] this DataTable dataTable) => ToExcelBytes(dataTable, ExcelFormat.Xls);

        /// <summary>
        ///     DataTable2ExcelBytes
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="excelFormat">excel格式</param>
        public static byte[] ToExcelBytes([NotNull] this DataTable dataTable, ExcelFormat excelFormat) => ToExcelBytes(dataTable, excelFormat, null);

        /// <summary>
        ///     DataTable2ExcelBytes
        /// </summary>
        /// <param name="dataTable">dataTable</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="excelSetting">excelSetting</param>
        public static byte[] ToExcelBytes([NotNull] this DataTable dataTable, ExcelFormat excelFormat, ExcelSetting excelSetting)
        {
            var workbook = ExcelHelper.PrepareWorkbook(excelFormat, excelSetting);
            var sheet = workbook.CreateSheet(
                string.IsNullOrWhiteSpace(dataTable.TableName)
                ? "Sheet0"
                : dataTable.TableName);
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

            return workbook.ToExcelBytes();
        }

        /// <summary>
        ///     SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        public static void SetCellValue([NotNull] this ICell cell, object value) => cell.SetCellValue(value, null);

        /// <summary>
        ///     SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        /// <param name="formatter">formatter</param>
        public static void SetCellValue([NotNull] this ICell cell, object value, string formatter)
        {
            if (null == value)
            {
                cell.SetCellType(CellType.Blank);
                return;
            }

            if (value is DateTime time)
            {
                cell.SetCellValue(string.IsNullOrWhiteSpace(formatter)
                    ? (time.Date == time ? time.ToStandardDateString() : time.ToStandardTimeString())
                    : time.ToString(formatter));
                cell.SetCellType(CellType.String);
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
                    cell.SetCellValue(Convert.ToDouble(value));
                    cell.SetCellType(CellType.Numeric);
                }
                else if (type == typeof(bool))
                {
                    cell.SetCellValue((bool)value);
                    cell.SetCellType(CellType.Boolean);
                }
                else
                {
                    cell.SetCellValue(value is IFormattable val && formatter.IsNotNullOrWhiteSpace()
                        ? val.ToString(formatter, CultureInfo.CurrentCulture)
                        : value.ToString());
                    cell.SetCellType(CellType.String);
                }
            }
        }

        /// <summary>
        ///     GetCellValue
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="propertyType">propertyType</param>
        /// <returns>cellValue</returns>
        public static object GetCellValue([CanBeNull] this ICell cell, Type propertyType)
        {
            if (cell == null || cell.CellType == CellType.Blank || cell.CellType == CellType.Error)
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
                    return cell.NumericCellValue.ToOrDefault(propertyType);

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
                    return cell.BooleanCellValue.ToOrDefault(propertyType);

                default:
                    return cell.ToString().ToOrDefault(propertyType);
            }
        }

        /// <summary>
        ///     GetCellValue
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="cell">cell</param>
        /// <returns></returns>
        public static T GetCellValue<T>([CanBeNull] this ICell cell) => (T)cell.GetCellValue(typeof(T));

        /// <summary>
        /// GetRowCollection
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <returns>row collection</returns>
        public static NpoiRowCollection GetRowCollection([NotNull]this ISheet sheet) => new NpoiRowCollection(sheet);

        /// <summary>
        ///     Write workbook to excel file
        /// </summary>
        /// <param name="workbook">workbook</param>
        /// <param name="filePath">file path</param>
        public static void WriteToFile([NotNull] this IWorkbook workbook, string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (null == dir)
            {
                filePath = ApplicationHelper.MapPath(filePath);
            }
            else
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }

            using (var fileStream = File.Create(filePath))
            {
                workbook.Write(fileStream);
            }
        }

        /// <summary>
        ///     ToExcelBytes
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
