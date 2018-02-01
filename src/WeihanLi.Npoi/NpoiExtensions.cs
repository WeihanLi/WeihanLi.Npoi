using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi
{
    public static class NpoiExtensions
    {
        /// <summary>
        /// Workbook2EntityList
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="workbook">excel workbook</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>entity list</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]this IWorkbook workbook, int sheetIndex = 0) where TEntity : new()
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
        public static List<TEntity> ToEntityList<TEntity>([NotNull]this ISheet sheet) where TEntity : new()
        {
            return new NpoiHelper<TEntity>().SheetToEntityList(sheet);
        }

        /// <summary>
        /// import entityList to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="list">entityList</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static void ImportData<TEntity>([NotNull]this IWorkbook workbook, IReadOnlyList<TEntity> list, int sheetIndex = 0) where TEntity : new()
        {
            if (sheetIndex >= ExcelConstants.MaxSheetNum)
            {
                throw new ArgumentException(string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), ExcelConstants.MaxSheetNum), nameof(sheetIndex));
            }

            while (workbook.NumberOfSheets <= sheetIndex)
            {
                workbook.CreateSheet();
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            new NpoiHelper<TEntity>().EntityListToSheet(sheet, list);
        }

        /// <summary>
        /// import entityList to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="list">entityList</param>
        public static void ImportData<TEntity>([NotNull]this ISheet sheet, IReadOnlyList<TEntity> list)
            where TEntity : new()
        {
            new NpoiHelper<TEntity>().EntityListToSheet(sheet, list);
        }

        /// <summary>
        /// import datatable to workbook first sheet
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="workbook">workbook</param>
        /// <param name="dataTable">dataTable</param>
        /// <param name="sheetIndex">sheetIndex</param>
        public static void ImportData<TEntity>([NotNull]this IWorkbook workbook, DataTable dataTable, int sheetIndex = 0) where TEntity : new()
        {
            if (sheetIndex >= ExcelConstants.MaxSheetNum)
            {
                throw new ArgumentException(string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), ExcelConstants.MaxSheetNum), nameof(sheetIndex));
            }

            while (workbook.NumberOfSheets <= sheetIndex)
            {
                workbook.CreateSheet();
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            new NpoiHelper<TEntity>().DataTableToSheet(sheet, dataTable);
        }

        /// <summary>
        /// import datatable to sheet
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="sheet">sheet</param>
        /// <param name="dataTable">dataTable</param>
        public static void ImportData<TEntity>([NotNull]this ISheet sheet, DataTable dataTable)
            where TEntity : new()
        {
            new NpoiHelper<TEntity>().DataTableToSheet(sheet, dataTable);
        }

        /// <summary>
        /// EntityList2ExcelFile
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="entityList">entityList</param>
        /// <param name="excelPath">excelPath</param>
        public static void ToExcelFile<TEntity>([NotNull] this IReadOnlyList<TEntity> entityList, string excelPath)
            where TEntity : new()
        {
            var workbook = ExcelHelper.PrepareWorkbook(excelPath);
            workbook.ImportData(entityList);
            workbook.WriteToFile(excelPath);
        }

        /// <summary>
        /// SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        public static void SetCellValue([NotNull]this ICell cell, object value)
        {
            if (null == value)
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue("");
                return;
            }
            if (value is DateTime time)
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(time.ToStandardTimeString());
            }
            var type = value.GetType().Unwrap();
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
                cell.SetCellValue(Convert.ToBoolean(value));
            }
            else
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(value.ToString());
            }
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
                        return cell.DateCellValue;
                    }
                    if (propertyType == typeof(double))
                    {
                        return cell.NumericCellValue;
                    }
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
            => cell.ToString().ToOrDefault<T>();

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
    }
}
