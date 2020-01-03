using System;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    public static class ConfigurationExtensions
    {
        /// <summary>
        /// get excelConfiguration property configuration
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="excelConfiguration">excelConfiguration</param>
        /// <param name="propertyName">propertyName</param>
        /// <returns>PropertyConfiguration</returns>
        public static IPropertyConfiguration<TEntity, string> Property<TEntity>(
            this IExcelConfiguration<TEntity> excelConfiguration, string propertyName) =>
            excelConfiguration.Property<string>(propertyName);

        /// <summary>
        /// has column output formatter
        /// </summary>
        /// <typeparam name="TEntity">entity type</typeparam>
        /// <typeparam name="TProperty">property type</typeparam>
        /// <param name="configuration">property configuration</param>
        /// <param name="formatter">column output formatter</param>
        /// <returns>property configuration</returns>
        public static IPropertyConfiguration<TEntity, TProperty> HasColumnOutputFormatter<TEntity, TProperty>(
            this IPropertyConfiguration<TEntity, TProperty> configuration, Func<TProperty, object> formatter) =>
            configuration.HasOutputFormatter((entity, prop) => formatter?.Invoke(prop));
    }
}
