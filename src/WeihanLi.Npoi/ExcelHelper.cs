using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using JetBrains.Annotations;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    /// <summary>
    /// ExcelHelper
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// 验证excel文件路径是否可用
        /// </summary>
        /// <param name="excelPath">路径</param>
        /// <param name="msg">错误信息</param>
        /// <param name="isExport">是否是导出操作，导出不需要验证是否存在</param>
        /// <returns>是否可用</returns>
        private static bool ValidateExcelFilePath(string excelPath, out string msg, bool isExport = false)
        {
            if (isExport || File.Exists(excelPath))
            {
                var ext = Path.GetExtension(excelPath);
                if (ext.EqualsIgnoreCase(".xls") || ext.EqualsIgnoreCase(".xlsx"))
                {
                    msg = string.Empty;
                    return true;
                }
                msg = Resource.InvalidFile;
                return false;
            }

            msg = Resource.FileNotFound;
            return false;
        }

        /// <summary>
        /// 根据excel路径加载excel
        /// </summary>
        /// <param name="excelPath">excel路径</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel(string excelPath)
        {
            if (!ValidateExcelFilePath(excelPath, out var msg))
                throw new ArgumentException(msg);

            using (var stream = File.OpenRead(excelPath))
            {
                return Path.GetExtension(excelPath).EqualsIgnoreCase(".xlsx") ? (IWorkbook)new XSSFWorkbook(stream) : new HSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// 为导出准备 workbook
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(string excelPath)
        {
            if (!ValidateExcelFilePath(excelPath, out var msg, true))
                throw new ArgumentException(msg);

            return Path.GetExtension(excelPath).EqualsIgnoreCase(".xlsx")
                ? new XSSFWorkbook() :
                (IWorkbook)new HSSFWorkbook();
        }

        /// <summary>
        /// 读取Excel内容到一个List中
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex，默认是0</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]string excelPath, int sheetIndex = 0) where TEntity : new()
        {
            var workbook = LoadExcel(excelPath);
            if (sheetIndex >= workbook.NumberOfSheets)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            return workbook.GetSheetAt(sheetIndex).ToEntityList<TEntity>();
        }

        /// <summary>
        /// 读取Excel内容到DataTable中
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex，默认是0</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable<TEntity>(string excelPath, int sheetIndex = 0) where TEntity : new()
            => ToEntityList<TEntity>(excelPath, sheetIndex).ToDataTable();

        /// <summary>
        /// 导出到excel中
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="entityList">entityList</param>
        /// <returns>export result</returns>
        public static int ExportToExcel<TEntity>(string excelPath, IReadOnlyList<TEntity> entityList) where TEntity : new()
        {
            // prepare workbook
            var workbook = PrepareWorkbook(excelPath);
            // import data
            ImportData(entityList, workbook);
            // save to file
            workbook.WriteToFile(excelPath);
            return 1;
        }

        /// <summary>
        /// 导出到excel中
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="dataTable">dataTable</param>
        /// <returns>export result</returns>
        public static int ExportToExcel<TEntity>(string excelPath, DataTable dataTable)
            where TEntity : new()
        {
            // prepare workbook
            var workbook = PrepareWorkbook(excelPath);
            // import data
            ImportData<TEntity>(dataTable, workbook);
            // save to file
            workbook.WriteToFile(excelPath);
            return 1;
        }

        public static void ImportData<TEntity>(DataTable dataTable, IWorkbook workbook, int sheetIndex = 0)
            where TEntity : new()
        {
            var sheet = typeof(TEntity).GetCustomAttribute<SheetAttribute>();
            if (null != sheet && sheet.SheetName.IsNotNullOrWhiteSpace())
            {
                while (sheetIndex >= workbook.NumberOfSheets && sheetIndex < ExcelConstants.MaxSheetNum)
                {
                    workbook.CreateSheet(sheet.SheetName);
                }
            }
            else
            {
                while (sheetIndex >= workbook.NumberOfSheets && sheetIndex < ExcelConstants.MaxSheetNum)
                {
                    workbook.CreateSheet();
                }
            }
            new NpoiHelper<TEntity>().DataTableToSheet(workbook.GetSheetAt(sheetIndex), dataTable);
        }

        public static void ImportData<TEntity>(IReadOnlyList<TEntity> list, IWorkbook workbook, int sheetIndex = 0)
            where TEntity : new()
        {
            while (sheetIndex >= workbook.NumberOfSheets && sheetIndex < ExcelConstants.MaxSheetNum)
            {
                workbook.CreateSheet();
            }
            new NpoiHelper<TEntity>().EntityListToSheet(workbook.GetSheetAt(sheetIndex), list);
        }
    }
}
