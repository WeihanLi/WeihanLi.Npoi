using System;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    public static class ConfigurationExtensions
    {
        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="configuration">excel configuration</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="sheetName">sheetName</param>
        /// <returns>current excel configuration</returns>
        public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
            string sheetName) => configuration.HasSheetConfiguration(config =>
                {
                    config.SheetName = sheetName;
                }, sheetIndex);

        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="configuration">excel configuration</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="sheetName">sheetName</param>
        /// <param name="enableAutoColumnWidth">enable auto column width if true otherwise false</param>
        /// <returns>current excel configuration</returns>
        public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
            string sheetName, bool enableAutoColumnWidth) => configuration.HasSheetConfiguration(config =>
        {
            config.SheetName = sheetName;
            config.AutoColumnWidthEnabled = enableAutoColumnWidth;
        }, sheetIndex);

        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="configuration">excel configuration</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="sheetName">sheetName</param>
        /// <param name="startRowIndex">startRowIndex</param>
        /// <returns>current excel configuration</returns>
        public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
            string sheetName, int startRowIndex) => configuration.HasSheetConfiguration(config =>
        {
            config.SheetName = sheetName;
            config.StartRowIndex = startRowIndex;
        }, sheetIndex);

        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="configuration">excel configuration</param>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="sheetName">sheetName</param>
        /// <param name="startRowIndex">startRowIndex</param>
        /// <param name="enableAutoColumnWidth">enable auto column width if true otherwise false</param>
        /// <param name="endRowIndex">endRowIndex, set this if you wanna control where to end(included)</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex, string sheetName, int startRowIndex,
            bool enableAutoColumnWidth, int? endRowIndex = null) => configuration.HasSheetConfiguration(config =>
        {
            config.SheetName = sheetName;
            config.StartRowIndex = startRowIndex;
            config.AutoColumnWidthEnabled = enableAutoColumnWidth;
            config.EndRowIndex = endRowIndex;
        }, sheetIndex);

        /// <summary>
        /// property configuration
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
            this IPropertyConfiguration<TEntity, TProperty> configuration, Func<TProperty, object> formatter)
        {
            if (formatter == null)
            {
                return configuration.HasOutputFormatter(null);
            }
            return configuration.HasOutputFormatter((entity, prop) => formatter.Invoke(prop));
        }
    }
}
