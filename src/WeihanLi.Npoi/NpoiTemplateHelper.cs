using JetBrains.Annotations;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi
{
    internal static class NpoiTemplateHelper
    {
        private const string GlobalParamFormat = "$(Global:{0})";

        private const string HeaderParamFormat = "$(Header:{0})";

        private const string DataBegin = "<Data>";
        private const string DataEnd = "</Data>";
        private const string DataParamFormat = "$(Data:{0})";

        // export via template
        public static ISheet EntityListToSheetByTemplate<TEntity>(
            [NotNull]ISheet sheet,
            IEnumerable<TEntity> entityList,
            object extraData = null)
        {
            if (null == entityList)
            {
                return sheet;
            }
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
            var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);

            var globalDictionary = extraData.ParseParamInfo()
                .ToDictionary(x => GlobalParamFormat.FormatWith(x.Key), x => x.Value);
            foreach (var propertyConfiguration in propertyColumnDictionary)
            {
                globalDictionary.Add(HeaderParamFormat.FormatWith(propertyConfiguration.Key.Name), propertyConfiguration.Value.ColumnTitle);
            }

            var dataFuncDictionary = propertyColumnDictionary
                .ToDictionary(x => DataParamFormat.FormatWith(x.Key.Name), x => x.Key.GetValueGetter());

            // parseTemplate
            int dataStartRow = -1, dataRowsCount = -1;

            for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }
                for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    var cellValue = cell.GetCellValue<string>();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        var beforeValue = cellValue;
                        if (dataStartRow <= 0 || dataRowsCount <= 0)
                        {
                            if (dataStartRow >= 0)
                            {
                                if (cellValue.Contains(DataEnd))
                                {
                                    dataRowsCount = rowIndex - dataStartRow + 1;
                                    cellValue = cellValue.Replace(DataEnd, string.Empty);
                                }
                            }
                            else
                            {
                                if (cellValue.Contains(DataBegin))
                                {
                                    dataStartRow = rowIndex;
                                    cellValue = cellValue.Replace(DataBegin, string.Empty);
                                }
                            }
                        }

                        foreach (var param in globalDictionary.Keys)
                        {
                            if (cellValue.Contains(param))
                            {
                                cellValue = cellValue
                                    .Replace(param,
                                    globalDictionary[param]?.ToString() ?? string.Empty);
                            }
                        }

                        if (beforeValue != cellValue)
                        {
                            cell.SetCellValue(cellValue);
                        }
                    }
                }
            }

            if (dataStartRow >= 0 && dataRowsCount > 0)
            {
                foreach (var entity in entityList)
                {
                    sheet.ShiftRows(dataStartRow, sheet.LastRowNum, dataRowsCount);
                    for (var i = 0; i < dataRowsCount; i++)
                    {
                        sheet.CopyRow(dataStartRow + dataRowsCount + i, dataStartRow + i);
                        var row = sheet.GetRow(dataStartRow + i);
                        if (null != row)
                        {
                            for (var j = 0; j < row.LastCellNum; j++)
                            {
                                var cell = row.GetCell(j);
                                if (null != cell)
                                {
                                    var cellValue = cell.GetCellValue<string>();
                                    if (!string.IsNullOrEmpty(cellValue))
                                    {
                                        var beforeValue = cellValue;
                                        foreach (var param in dataFuncDictionary.Keys)
                                        {
                                            if (cellValue.Contains(param))
                                            {
                                                cellValue = cellValue.Replace(param,
                                                    dataFuncDictionary[param]?.Invoke(entity)?.ToString() ?? string.Empty);
                                            }
                                        }

                                        if (beforeValue != cellValue)
                                        {
                                            cell.SetCellValue(cellValue);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //
                    dataStartRow += dataRowsCount;
                }

                // remove template
                for (var i = 0; i < dataRowsCount; i++)
                {
                    var row = sheet.GetRow(dataStartRow + i);
                    if (null != row)
                    {
                        sheet.RemoveRow(row);
                    }
                }

                sheet.ShiftRows(dataStartRow + dataRowsCount, sheet.LastRowNum, -dataRowsCount);
            }

            return sheet;
        }
    }
}
