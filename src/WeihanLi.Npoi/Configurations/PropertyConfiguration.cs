using System;
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
    }

    internal class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration, IPropertyConfiguration<TEntity, TProperty>
    {
        /// <summary>
        /// InputFormatterFunc
        /// </summary>
        public Func<TEntity, TProperty, TProperty> InputFormatterFunc { get; set; }

        /// <summary>
        /// OutputFormatterFunc
        /// </summary>
        public Func<TEntity, TProperty, object> OutputFormatterFunc { get; set; }

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

        public IPropertyConfiguration<TEntity, TProperty> Ignored()
        {
            IsIgnored = true;
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            if (formatterFunc != null)
            {
                OutputFormatterFunc = formatterFunc;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            if (formatterFunc != null)
            {
                OutputFormatterFunc = formatterFunc;
            }
            return this;
        }

        public IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(
            Func<TEntity, TProperty, TProperty> formatterFunc)
        {
            if (formatterFunc != null)
            {
                InputFormatterFunc = formatterFunc;
            }
            return this;
        }
    }
}
