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
                FilterSetting =
                    type.GetCustomAttribute<FilterAttribute>()?.FilterSeting,
                FreezeSettings =
                    type.GetCustomAttributes<FreezeAttribute>().Select(_ => _.FreezeSetting).ToList()
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
                dic.Add(propertyInfo, new PropertyConfiguration(column.PropertySetting));
            }
            excelConfiguration.PropertyConfigurationDictionary = dic;
            return excelConfiguration;
        }
    }
}
