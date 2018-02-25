using System;
using System.Collections.Concurrent;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    internal static class TypeCache
    {
        internal static readonly ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConfigurationDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();
    }
}
