using System;
using System.Collections.Concurrent;
using System.Reflection;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    internal static class InternalCache
    {
        /// <summary>
        /// TypeExcelConfigurationCache
        /// </summary>
        internal static readonly ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConfigurationDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();

        internal static readonly ConcurrentDictionary<PropertyInfo, object> ColumnFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, object>();
    }
}
