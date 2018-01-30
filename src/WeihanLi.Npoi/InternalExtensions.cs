using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    internal static class InternalExtensions
    {
        #region Mapping

        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Keys.ToArray()[index];

        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Values.ToArray()[index];

        /// <summary>
        /// 根据属性名称获取属性信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="propertyName">属性名称</param>
        /// <returns></returns>
        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string propertyName) => mappingDictionary.Keys.FirstOrDefault(k => k.Name.EqualsIgnoreCase(propertyName));

        /// <summary>
        /// 根据属性名称获取列信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="propertyName">属性名称</param>
        /// <returns></returns>
        internal static ColumnAttribute GetColumnAttributeByPropertyName([NotNull] this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Keys.Any(_ => _.Name.EqualsIgnoreCase(propertyName)) ?
                mappingDictionary[mappingDictionary.Keys.First(_ => _.Name.EqualsIgnoreCase(propertyName))] : null;

        /// <summary>
        /// 根据列显示名称获取列信息
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="columnTitle">列名称</param>
        /// <returns></returns>
        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.Title.EqualsIgnoreCase(columnTitle));

        #endregion Mapping
    }
}
