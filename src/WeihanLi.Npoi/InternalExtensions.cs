using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal static class InternalExtensions
    {
        /// <summary>
        /// 根据属性名称获取列信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="propertyName">属性名称</param>
        /// <returns></returns>
        internal static PropertySetting GetPropertySettingByPropertyName([NotNull] this IDictionary<PropertyInfo, PropertySetting> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Keys.Any(_ => _.Name.EqualsIgnoreCase(propertyName)) ?
                mappingDictionary[mappingDictionary.Keys.First(_ => _.Name.EqualsIgnoreCase(propertyName))] : null;

        /// <summary>
        /// 根据列显示名称获取列信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="columnTitle">列名称</param>
        /// <returns></returns>
        internal static PropertySetting GetPropertySetting([NotNull]this IDictionary<PropertyInfo, PropertySetting> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.ColumnTitle.EqualsIgnoreCase(columnTitle));

        /// <summary>
        /// 根据列显示名称获取列信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="columnTitle">列名称</param>
        /// <returns></returns>
        internal static PropertySetting<TProperty> GetPropertySetting<TProperty>([NotNull]this IDictionary<PropertyInfo, PropertySetting> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.ColumnTitle.EqualsIgnoreCase(columnTitle)) as PropertySetting<TProperty>;
    }
}
