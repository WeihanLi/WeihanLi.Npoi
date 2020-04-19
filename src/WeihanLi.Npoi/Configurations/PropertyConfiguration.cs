using System;
using System.Reflection;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi.Configurations
{
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        /// <summary>
        /// ColumnIndex
        /// </summary>
        public int ColumnIndex { get; set; } = -1;

        /// <summary>
        /// ColumnWidth
        /// </summary>
        public int ColumnWidth { get; set; }

        /// <summary>
        /// Title
        /// </summary>
        public string ColumnTitle { get; set; }

        /// <summary>
        /// Formatter
        /// </summary>
        public string ColumnFormatter { get; set; }

        /// <summary>
        /// the property is ignored.
        /// </summary>
        public bool IsIgnored { get; set; }

        /// <summary>
        /// PropertyName
        /// </summary>
        public string PropertyName { get; set; }
    }

    internal sealed class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration, IPropertyConfiguration<TEntity, TProperty>
    {
        private readonly PropertyInfo _propertyInfo;

        public PropertyConfiguration(PropertyInfo propertyInfo)
        {
            _propertyInfo = propertyInfo;
            PropertyName = propertyInfo.Name;
            ColumnTitle = propertyInfo.Name;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index)
        {
            if (index >= 0)
            {
                ColumnIndex = index;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title)
        {
            if (title.IsNotNullOrWhiteSpace())
            {
                ColumnTitle = title;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width)
        {
            ColumnWidth = width;
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter)
        {
            if (formatter.IsNotNullOrWhiteSpace())
            {
                ColumnFormatter = formatter;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> Ignored(bool ignored = true)
        {
            IsIgnored = ignored;
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            InternalCache.OutputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(
            Func<TEntity, TProperty, TProperty> formatterFunc)
        {
            InternalCache.InputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc)
        {
            InternalCache.ColumnInputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
            return this;
        }
    }
}
