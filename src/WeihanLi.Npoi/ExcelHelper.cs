using JetBrains.Annotations;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
        /// Default excel setting for export excel files
        /// </summary>
        public static ExcelSetting DefaultExcelSetting = new ExcelSetting();

        /// <summary>
        /// Validate is a excel path valid
        /// </summary>
        /// <param name="excelPath">excel path</param>
        /// <param name="msg">error message</param>
        /// <param name="isExport">is export operation</param>
        /// <returns>is valid excel path</returns>
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
                msg = Resource.InvalidExcelFile;
                return false;
            }

            msg = Resource.FileNotFound;
            return false;
        }

        /// <summary>
        /// load excel from filepath
        /// </summary>
        /// <param name="excelPath">excel file path</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel([NotNull]string excelPath)
        {
            if (!ValidateExcelFilePath(excelPath, out var msg))
            {
                throw new ArgumentException(msg);
            }

            using (var stream = File.OpenRead(excelPath))
            {
                return Path.GetExtension(excelPath).EqualsIgnoreCase(".xls") ? new HSSFWorkbook(stream) : (IWorkbook)new XSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// load excel from excelBytes
        /// </summary>
        /// <param name="excelBytes">excel file bytes</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel([NotNull]byte[] excelBytes) => LoadExcel(excelBytes, ExcelFormat.Xls);

        /// <summary>
        /// load excel from excelBytes
        /// </summary>
        /// <param name="excelBytes">excel file bytes</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel([NotNull]byte[] excelBytes, ExcelFormat excelFormat)
        {
            using (var stream = new MemoryStream(excelBytes))
            {
                return LoadExcel(stream, excelFormat);
            }
        }

        /// <summary>
        /// load excel from excelBytes
        /// </summary>
        /// <param name="excelStream">excel file stream</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel([NotNull]Stream excelStream) => LoadExcel(excelStream, ExcelFormat.Xls);

        /// <summary>
        /// load excel from excelBytes
        /// </summary>
        /// <param name="excelStream">excel file stream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns>workbook</returns>
        public static IWorkbook LoadExcel([NotNull]Stream excelStream, ExcelFormat excelFormat)
        {
            switch (excelFormat)
            {
                case ExcelFormat.Xls:
                    return new HSSFWorkbook(excelStream);

                // case ExcelFormat.Xlsx:
                default:
                    return new XSSFWorkbook(excelStream);
            }
        }

        /// <summary>
        /// prepare a workbook for export
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook([NotNull]string excelPath) => PrepareWorkbook(excelPath, null);

        /// <summary>
        /// prepare a workbook for export
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <param name="excelSetting">excelSetting</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook([NotNull]string excelPath, ExcelSetting excelSetting)
        {
            if (!ValidateExcelFilePath(excelPath, out var msg, true))
            {
                throw new ArgumentException(msg);
            }
            return PrepareWorkbook(!Path.GetExtension(excelPath).EqualsIgnoreCase(".xls"), excelSetting);
        }

        /// <summary>
        /// prepare a workbook for export
        /// </summary>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="excelSetting">excelSetting</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(ExcelFormat excelFormat, ExcelSetting excelSetting)
        {
            return PrepareWorkbook(excelFormat == ExcelFormat.Xlsx, excelSetting);
        }

        /// <summary>
        /// get a excel workbook(*.xlsx)
        /// </summary>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook() => PrepareWorkbook(true);

        /// <summary>
        /// get a excel workbook
        /// </summary>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(ExcelFormat excelFormat) => PrepareWorkbook(excelFormat == ExcelFormat.Xlsx);

        /// <summary>
        /// get a excel workbook
        /// </summary>
        /// <param name="isXlsx">is for *.xlsx file</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(bool isXlsx) => PrepareWorkbook(isXlsx, null);

        private static readonly Version Version = typeof(ExcelHelper).Assembly.GetName().Version;

        /// <summary>
        /// get a excel workbook
        /// </summary>
        /// <param name="isXlsx">is for *.xlsx file</param>
        /// <param name="excelSetting">excelSettings</param>
        /// <returns></returns>
        public static IWorkbook PrepareWorkbook(bool isXlsx, ExcelSetting excelSetting)
        {
            var setting = (excelSetting ?? DefaultExcelSetting) ?? new ExcelSetting();

            if (isXlsx)
            {
                var workbook = new XSSFWorkbook();
                var props = workbook.GetProperties();
                props.CoreProperties.Creator = setting.Author;
                props.CoreProperties.Created = DateTime.Now;
                props.CoreProperties.Modified = DateTime.Now;
                props.CoreProperties.Title = setting.Title;
                props.CoreProperties.Subject = setting.Subject;
                props.CoreProperties.Category = setting.Category;
                props.CoreProperties.Description = setting.Description;
                props.ExtendedProperties.GetUnderlyingProperties().Company = setting.Company;
                props.ExtendedProperties.GetUnderlyingProperties().Application = InternalConstants.ApplicationName;
                props.ExtendedProperties.GetUnderlyingProperties().AppVersion = Version.ToString();
                return workbook;
            }
            else
            {
                var workbook = new HSSFWorkbook();
                ////create a entry of DocumentSummaryInformation
                var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = setting.Company;
                dsi.Category = setting.Category;
                workbook.DocumentSummaryInformation = dsi;
                ////create a entry of SummaryInformation
                var si = PropertySetFactory.CreateSummaryInformation();
                si.Title = setting.Title;
                si.Subject = setting.Subject;
                si.Author = setting.Author;
                si.CreateDateTime = DateTime.Now;
                si.Comments = setting.Description;
                si.ApplicationName = InternalConstants.ApplicationName;
                workbook.SummaryInformation = si;
                return workbook;
            }
        }

        /// <summary>
        /// read first sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelBytes">excelBytes</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]byte[] excelBytes) where TEntity : new()
            => ToEntityList<TEntity>(excelBytes, ExcelFormat.Xls, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelBytes">excelBytes</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] byte[] excelBytes, int sheetIndex)
            where TEntity : new()
            => ToEntityList<TEntity>(excelBytes, ExcelFormat.Xls, sheetIndex);

        /// <summary>
        /// read first sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelBytes">excelBytes</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] byte[] excelBytes, ExcelFormat excelFormat)
            where TEntity : new() => ToEntityList<TEntity>(excelBytes, excelFormat, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel bytes path to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelBytes">excelBytes</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]byte[] excelBytes, ExcelFormat excelFormat, int sheetIndex) where TEntity : new()
        {
            var workbook = LoadExcel(excelBytes, excelFormat);
            return workbook.ToEntityList<TEntity>(sheetIndex);
        }

        /// <summary>
        /// read first sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelStream">excelStream</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]Stream excelStream) where TEntity : new()
            => ToEntityList<TEntity>(excelStream, ExcelFormat.Xls, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelStream">excelStream</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] Stream excelStream, int sheetIndex)
            where TEntity : new()
            => ToEntityList<TEntity>(excelStream, ExcelFormat.Xls, sheetIndex);

        /// <summary>
        /// read first sheet of excel from excel file bytes to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelStream">excelStream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull] Stream excelStream, ExcelFormat excelFormat)
            where TEntity : new()
            => ToEntityList<TEntity>(excelStream, excelFormat, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel bytes path to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelStream">excelStream</param>
        /// <param name="excelFormat">excelFormat</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]Stream excelStream, ExcelFormat excelFormat, int sheetIndex) where TEntity : new()
        {
            var workbook = LoadExcel(excelStream, excelFormat);
            return workbook.ToEntityList<TEntity>(sheetIndex);
        }

        /// <summary>
        /// read first sheet of excel from excel file path to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]string excelPath) where TEntity : new() => ToEntityList<TEntity>(excelPath, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel file path to a list
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>List</returns>
        public static List<TEntity> ToEntityList<TEntity>([NotNull]string excelPath, int sheetIndex) where TEntity : new()
        {
            var workbook = LoadExcel(excelPath);
            return workbook.ToEntityList<TEntity>(sheetIndex);
        }

        /// <summary>
        /// read first sheet of excel from excel file path to a data table
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable<TEntity>([NotNull]string excelPath) where TEntity : new() => ToDataTable<TEntity>(excelPath, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel file path to a list(for specific class type)
        /// </summary>
        /// <typeparam name="TEntity">EntityType</typeparam>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable<TEntity>([NotNull]string excelPath, int sheetIndex) where TEntity : new()
            => ToEntityList<TEntity>(excelPath, sheetIndex).ToDataTable();

        /// <summary>
        /// read first sheet of excel from excel file path to a data table
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]string excelPath) => ToDataTable(excelPath, 0, 0);

        /// <summary>
        /// read (sheetIndex) sheet of excel from excel file path to a data table
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable([NotNull]string excelPath, int sheetIndex, int headerRowIndex)
        {
            var workbook = LoadExcel(excelPath);
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), string.Format(Resource.IndexOutOfRange, nameof(sheetIndex), workbook.NumberOfSheets));
            }
            return workbook.GetSheetAt(sheetIndex).ToDataTable();
        }

        /// <summary>
        /// read first sheet of excel from excel file path to a DataSet from second row
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <returns></returns>
        public static DataSet ToDataSet([NotNull]string excelPath) => ToDataSet(excelPath, 0);

        /// <summary>
        /// read first sheet of excel from excel file path to a DataSet from (headerRowIndex+1) row
        /// </summary>
        /// <param name="excelPath">excelPath</param>
        /// <param name="headerRowIndex">headerRowIndex</param>
        /// <returns></returns>
        public static DataSet ToDataSet([NotNull]string excelPath, int headerRowIndex) => LoadExcel(excelPath).ToDataSet(headerRowIndex);

        /// <summary>
        /// SettingFor
        /// Fluent Settings
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns></returns>
        public static IExcelConfiguration<TEntity> SettingFor<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();
    }
}
