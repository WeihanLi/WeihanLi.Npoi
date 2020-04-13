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
        public static readonly ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConfigurationDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();

        public static readonly ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>> OutputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>>();

        public static readonly ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>> InputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>>();

        public static readonly ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>> ColumnInputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Tuple<MethodInfo, object>>();
    }
}
