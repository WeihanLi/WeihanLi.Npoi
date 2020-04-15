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

        public static readonly ConcurrentDictionary<PropertyInfo, DelegateInvokeContext> OutputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, DelegateInvokeContext>();

        public static readonly ConcurrentDictionary<PropertyInfo, DelegateInvokeContext> InputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, DelegateInvokeContext>();

        public static readonly ConcurrentDictionary<PropertyInfo, DelegateInvokeContext> ColumnInputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, DelegateInvokeContext>();
    }
}
