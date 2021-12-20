using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

internal static class InternalHelper
{
    /// <summary>
    ///     Get ExcelConfigurationMapping by type
    /// </summary>
    /// <param name="entityType">entityType</param>
    /// <returns>excel configuration</returns>
    public static IExcelConfiguration GetExcelConfigurationMapping(Type entityType) =>
        InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(entityType, type =>
        {
            var excelConfiguration =
                CreateExcelConfiguration(type, () => (ExcelConfiguration)Guard.NotNull(Activator.CreateInstance(
                    typeof(ExcelConfiguration<>).MakeGenericType(entityType))));
            return excelConfiguration;
        });

    /// <summary>
    ///     Get GenericType ExcelConfigurationMapping
    /// </summary>
    /// <typeparam name="TEntity">TEntity</typeparam>
    /// <returns>IExcelConfiguration</returns>
    public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>() =>
        (ExcelConfiguration<TEntity>)InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(typeof(TEntity),
            type =>
            {
                var excelConfiguration =
                    CreateExcelConfiguration(type, () => new ExcelConfiguration<TEntity>());
                return excelConfiguration;
            });

    private static ExcelConfiguration CreateExcelConfiguration(Type type,
        Func<ExcelConfiguration> newConfigurationFunc)
    {
        var excelConfiguration = newConfigurationFunc();
        excelConfiguration.FilterSetting = type.GetCustomAttribute<FilterAttribute>()?.FilterSetting;
        foreach (var sheetAttribute in type.GetCustomAttributes<SheetAttribute>())
        {
            if (sheetAttribute.SheetIndex >= 0)
            {
                excelConfiguration.SheetSettings[sheetAttribute.SheetIndex] = sheetAttribute.SheetSetting;
            }
        }

        foreach (var freezeAttribute in type.GetCustomAttributes<FreezeAttribute>())
        {
            excelConfiguration.FreezeSettings.Add(freezeAttribute.FreezeSetting);
        }

        var propertyInfos = CacheUtil.GetTypeProperties(type);
        foreach (var propertyInfo in propertyInfos)
        {
            var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
            if (string.IsNullOrWhiteSpace(column.Title))
            {
                column.Title = propertyInfo.Name;
            }

            var propertyConfigurationType =
                typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
            var propertyConfiguration = Guard.NotNull(Activator.CreateInstance(propertyConfigurationType, propertyInfo));

            propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnTitle))
                ?.GetSetMethod()?
                .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnTitle });
            propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnIndex))
                ?.GetSetMethod()?
                .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnIndex });
            propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnFormatter))
                ?.GetSetMethod()?
                .Invoke(propertyConfiguration, new object?[] { column.PropertyConfiguration.ColumnFormatter });
            propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.IsIgnored))
                ?.GetSetMethod()?
                .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.IsIgnored });
            propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnWidth))
                ?.GetSetMethod()?
                .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnWidth });

            excelConfiguration.PropertyConfigurationDictionary.Add(propertyInfo,
                (PropertyConfiguration)propertyConfiguration);
        }

        return excelConfiguration;
    }

    /// <summary>
    ///     Adjust Column Index
    /// </summary>
    /// <typeparam name="TEntity">TEntity</typeparam>
    /// <param name="excelConfiguration">excelConfiguration</param>
    private static void AdjustColumnIndex<TEntity>(ExcelConfiguration<TEntity> excelConfiguration)
    {
        ICollection<int> validColumnIndex = excelConfiguration.PropertyConfigurationDictionary.Values
            .Where(c => !c.IsIgnored)
            .Select(x => x.ColumnIndex)
            .ToArray();
        // return if already adjusted
        if (validColumnIndex.All(_ => _ >= 0) &&
            validColumnIndex.Distinct().Count() == validColumnIndex.Count)
        {
            return;
        }

        var colIndexList = new List<int>(validColumnIndex.Count);
        foreach (var item in excelConfiguration.PropertyConfigurationDictionary
            .Where(_ => !_.Value.IsIgnored)
            .OrderBy(_ => _.Value.ColumnIndex >= 0 ? _.Value.ColumnIndex : int.MaxValue)
            .ThenBy(_ => _.Key.Name)
            .Select(_ => _.Value)
        )
        {
            while (colIndexList.Contains(item.ColumnIndex) || item.ColumnIndex < 0)
            {
                if (colIndexList.Count > 0)
                {
                    item.ColumnIndex = colIndexList.Max() + 1;
                }
                else
                {
                    item.ColumnIndex++;
                }
            }

            colIndexList.Add(item.ColumnIndex);
        }
    }

    /// <summary>
    ///     GetPropertyColumnDictionary
    /// </summary>
    /// <typeparam name="TEntity">TEntity Type</typeparam>
    /// <returns></returns>
    public static Dictionary<PropertyInfo, PropertyConfiguration> GetPropertyColumnDictionary<TEntity>() =>
        GetPropertyColumnDictionary(GetExcelConfigurationMapping<TEntity>());

    /// <summary>
    ///     GetPropertyColumnDictionary
    /// </summary>
    /// <typeparam name="TEntity">TEntity Type</typeparam>
    /// <returns></returns>
    public static Dictionary<PropertyInfo, PropertyConfiguration> GetPropertyColumnDictionary<TEntity>(
        ExcelConfiguration<TEntity> configuration)
    {
        AdjustColumnIndex(configuration);

        return configuration.PropertyConfigurationDictionary
            .Where(p => !p.Value.IsIgnored)
            .ToDictionary(_ => _.Key, _ => _.Value);
    }

    /// <summary>
    ///     GetProperties
    /// </summary>
    /// <typeparam name="TEntity">TEntity Type</typeparam>
    /// <returns></returns>
    public static IReadOnlyList<PropertyInfo> GetPropertiesForCsvHelper<TEntity>()
    {
        var configuration = GetExcelConfigurationMapping<TEntity>();

        AdjustColumnIndex(configuration);

        return configuration.PropertyConfigurationDictionary
            .Where(p => !p.Value.IsIgnored)
            .OrderBy(p => p.Value.ColumnIndex)
            .Select(p => p.Key)
            .ToArray();
    }

    public static string GetEncodedColumnName(string columnName) =>
        $"{columnName}{InternalConstants.DuplicateColumnMark}{Guid.NewGuid():N}";

    public static string GetDecodeColumnName(string columnName)
    {
        var duplicateMarkIndex = columnName.IndexOf(InternalConstants.DuplicateColumnMark, StringComparison.OrdinalIgnoreCase);
        if (duplicateMarkIndex > 0)
        {
            return columnName.Substring(0, duplicateMarkIndex);
        }
        return columnName;
    }
}
