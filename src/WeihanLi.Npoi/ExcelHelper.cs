using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using JetBrains.Annotations;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

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
        public static IWorkbook LoadExcel([NotNull]string excelPath)
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
        public static IWorkbook PrepareWorkbook([NotNull] string excelPath) => PrepareWorkbook(excelPath, null);

        /// <summary>
        /// 为导出准备 workbook
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <param name="excelSetting">excelSetting</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(string excelPath, ExcelSetting excelSetting)
        {
            if (!ValidateExcelFilePath(excelPath, out var msg, true))
                throw new ArgumentException(msg);
            return PrepareWorkbook(Path.GetExtension(excelPath).EqualsIgnoreCase(".xlsx"), excelSetting);
        }

        /// <summary>
        /// 获取一个 Excel Workbook（xlsx格式）
        /// </summary>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook() => PrepareWorkbook(true);

        /// <summary>
        /// 获取一个Excel workbook
        /// </summary>
        /// <param name="isXlsx">是否是 Xlsx 格式</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(bool isXlsx) =>
            PrepareWorkbook(isXlsx, null);

        /// <summary>
        /// 获取一个Excel workbook
        /// </summary>
        /// <param name="isXlsx">是否是 Xlsx 格式</param>
        /// <param name="excelSetting">excelSettings</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(bool isXlsx, ExcelSetting excelSetting)
        {
            if (null == excelSetting)
            {
                excelSetting = new ExcelSetting();
            }
            if (isXlsx)
            {
                var workbook = new XSSFWorkbook();
                var props = workbook.GetProperties();
                props.CoreProperties.Creator = excelSetting.Author;
                props.CoreProperties.Created = DateTime.Now;
                props.CoreProperties.Title = excelSetting.Title;
                props.CoreProperties.Subject = excelSetting.Subject;
                props.CoreProperties.Description = excelSetting.Description;
                props.ExtendedProperties.GetUnderlyingProperties().Application = ExcelConstants.ApplicationName;
                return workbook;
            }
            else
            {
                var workbook = new HSSFWorkbook();
                ////create a entry of DocumentSummaryInformation
                var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = ExcelConstants.ApplicationName;
                workbook.DocumentSummaryInformation = dsi;
                ////create a entry of SummaryInformation
                var si = PropertySetFactory.CreateSummaryInformation();
                si.Title = excelSetting.Title;
                si.Subject = excelSetting.Subject;
                si.Author = excelSetting.Author;
                si.CreateDateTime = DateTime.Now;
                si.Comments = excelSetting.Description;
                si.ApplicationName = ExcelConstants.ApplicationName;
                workbook.SummaryInformation = si;
                return workbook;
            }
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
            return workbook.ToEntityList<TEntity>(sheetIndex);
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
        /// 读取Excel内容到DataTable中
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex，默认是0</param>
        /// <param name="headerRowIndex">列首行 headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable(string excelPath, int sheetIndex = 0, int headerRowIndex = 0)
        {
            var workbook = LoadExcel(excelPath);
            var dataTable = new DataTable();
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            var rowEnumerator = sheet.GetRowEnumerator();
            while (rowEnumerator.MoveNext())
            {
                var row = (IRow)rowEnumerator.Current;
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

        #region FluentAPI

        /// <summary>
        /// FluentAPI
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns></returns>
        public static ExcelConfiguration<TEntity> SettingFor<TEntity>()
        {
            return new ExcelConfiguration<TEntity>();
        }

        /// <summary>
        /// FluentAPI
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="setting">ExcelSetting</param>
        /// <returns></returns>
        public static ExcelConfiguration<TEntity> SettingFor<TEntity>(ExcelSetting setting)
        {
            return new ExcelConfiguration<TEntity>(setting);
        }

        #endregion FluentAPI
    }
}
