using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal static class InternalHelper
    {
        /// <summary>
        /// GetExcelConfigurationMapping
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns>IExcelConfiguration</returns>
        public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>()
        {
            var type = typeof(TEntity);
            var excelConfiguration = new ExcelConfiguration<TEntity>
            {
                SheetSettings = new[]
                {
                    type.GetCustomAttribute<SheetAttribute>()?.SheetSetting?? new SheetSetting()
                },
                FilterSetting = type.GetCustomAttribute<FilterAttribute>()?.FilterSetting,
                FreezeSettings = type.GetCustomAttributes<FreezeAttribute>().Select(_ => _.FreezeSetting).ToList()
            };

            // propertyInfos
            var dic = new Dictionary<PropertyInfo, PropertyConfiguration>();
            var propertyInfos = Common.CacheUtil.TypePropertyCache.GetOrAdd(type, t => t.GetProperties());
            foreach (var propertyInfo in propertyInfos)
            {
                var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                if (string.IsNullOrWhiteSpace(column.Title))
                {
                    column.Title = propertyInfo.Name;
                }

                var propertySettingType = typeof(PropertySetting<,>).MakeGenericType(type, propertyInfo.PropertyType);
                var propertySetting = Activator.CreateInstance(propertySettingType);

                propertySettingType.GetProperty("ColumnTitle")?.GetSetMethod()?.Invoke(propertySetting, new object[] { column.PropertySetting.ColumnTitle });
                propertySettingType.GetProperty("ColumnIndex")?.GetSetMethod()?.Invoke(propertySetting, new object[] { column.PropertySetting.ColumnIndex });
                propertySettingType.GetProperty("ColumnFormatter")?.GetSetMethod()?.Invoke(propertySetting, new object[] { column.PropertySetting.ColumnFormatter });
                propertySettingType.GetProperty("IsIgnored")?.GetSetMethod()?.Invoke(propertySetting, new object[] { column.PropertySetting.IsIgnored });

                var propertyConfigurationType =
                    typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
                var propertyConfiguration = Activator.CreateInstance(propertyConfigurationType, new object[] { propertySetting });

                dic.Add(propertyInfo, (PropertyConfiguration)propertyConfiguration);
            }
            excelConfiguration.PropertyConfigurationDictionary = dic;

            return excelConfiguration;
        }

        public static void AdjustColumnIndex<TEntity>(ExcelConfiguration<TEntity> excelConfiguration)
        {
            var colIndexList = new List<int>(excelConfiguration.PropertyConfigurationDictionary.Count);

            foreach (var item in excelConfiguration.PropertyConfigurationDictionary
                .Where(_ => !_.Value.PropertySetting.IsIgnored)
                .OrderBy(_ => _.Value.PropertySetting.ColumnIndex >= 0 ? _.Value.PropertySetting.ColumnIndex : int.MaxValue)
                .ThenBy(_ => _.Key.Name)
                .Select(_ => _.Value)
                )
            {
                while (colIndexList.Contains(item.PropertySetting.ColumnIndex) || item.PropertySetting.ColumnIndex < 0)
                {
                    if (colIndexList.Count > 0)
                    {
                        item.PropertySetting.ColumnIndex = colIndexList.Max() + 1;
                    }
                    else
                    {
                        item.PropertySetting.ColumnIndex++;
                    }
                }
                colIndexList.Add(item.PropertySetting.ColumnIndex);
            }
        }

        /// <summary>
        /// GetProperties
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static PropertyInfo[] GetPropertiesForCsvHelper<TEntity>()
        {
            var configuration = (ExcelConfiguration<TEntity>)InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(typeof(TEntity), t => GetExcelConfigurationMapping<TEntity>());

            AdjustColumnIndex(configuration);

            return configuration.PropertyConfigurationDictionary
                .Where(p => !p.Value.PropertySetting.IsIgnored)
                .OrderBy(p => p.Value.PropertySetting.ColumnIndex)
                .ThenBy(p => p.Key.Name)
                .Select(p => p.Key)
                .ToArray();
        }
    }
}
