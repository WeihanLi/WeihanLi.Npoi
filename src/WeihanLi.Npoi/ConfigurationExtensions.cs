// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

public static class ConfigurationExtensions
{
    /// <summary>
    ///     Sheet Configuration
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <param name="sheetName">sheetName</param>
    /// <returns>current excel configuration</returns>
    public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
        string sheetName) => configuration.HasSheetSetting(config => { config.SheetName = sheetName; }, sheetIndex);

    /// <summary>
    ///     Sheet Configuration
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <param name="sheetName">sheetName</param>
    /// <param name="enableAutoColumnWidth">enable auto column width if true otherwise false</param>
    /// <returns>current excel configuration</returns>
    public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
        string sheetName, bool enableAutoColumnWidth) => configuration.HasSheetSetting(config =>
    {
        config.SheetName = sheetName;
        config.AutoColumnWidthEnabled = enableAutoColumnWidth;
    }, sheetIndex);

    /// <summary>
    ///     Sheet Configuration
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <param name="sheetName">sheetName</param>
    /// <param name="startRowIndex">startRowIndex</param>
    /// <returns>current excel configuration</returns>
    public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
        string sheetName, int startRowIndex) => configuration.HasSheetSetting(config =>
    {
        config.SheetName = sheetName;
        config.StartRowIndex = startRowIndex;
    }, sheetIndex);

    /// <summary>
    ///     Sheet Configuration
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <param name="sheetName">sheetName</param>
    /// <param name="startRowIndex">startRowIndex</param>
    /// <param name="enableAutoColumnWidth">enable auto column width if true otherwise false</param>
    /// <param name="endRowIndex">endRowIndex, set this if you wanna control where to end(included)</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasSheetConfiguration(this IExcelConfiguration configuration, int sheetIndex,
        string sheetName, int startRowIndex,
        bool enableAutoColumnWidth, int? endRowIndex = null) => configuration.HasSheetSetting(config =>
    {
        config.SheetName = sheetName;
        config.StartRowIndex = startRowIndex;
        config.AutoColumnWidthEnabled = enableAutoColumnWidth;
        config.EndRowIndex = endRowIndex;
    }, sheetIndex);

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="author">excel document author name</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasAuthor(this IExcelConfiguration configuration, string author) =>
        configuration.HasExcelSetting(setting => { setting.Author = author; });

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="title">excel document title</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasTitle(this IExcelConfiguration configuration, string title) =>
        configuration.HasExcelSetting(setting => { setting.Title = title; });

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="description">excel document description</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasDescription(this IExcelConfiguration configuration, string description) =>
        configuration.HasExcelSetting(setting => { setting.Description = description; });

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="subject">excel document subject</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasSubject(this IExcelConfiguration configuration, string subject) =>
        configuration.HasExcelSetting(setting => { setting.Subject = subject; });

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="company">excel document company</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasCompany(this IExcelConfiguration configuration, string company) =>
        configuration.HasExcelSetting(setting => { setting.Company = company; });

    /// <summary>
    ///     Configure excel author
    /// </summary>
    /// <param name="configuration">excel configuration</param>
    /// <param name="category">excel document category</param>
    /// <returns>current excel configuration<see cref="IExcelConfiguration" /></returns>
    public static IExcelConfiguration HasCategory(this IExcelConfiguration configuration, string category) =>
        configuration.HasExcelSetting(setting => { setting.Category = category; });

    /// <summary>
    ///     property configuration
    /// </summary>
    /// <typeparam name="TEntity">TEntity</typeparam>
    /// <param name="excelConfiguration">excelConfiguration</param>
    /// <param name="propertyName">propertyName</param>
    /// <returns>PropertyConfiguration</returns>
    public static IPropertyConfiguration<TEntity, string> Property<TEntity>(
        this IExcelConfiguration<TEntity> excelConfiguration, string propertyName) =>
        excelConfiguration.Property<string>(propertyName);

    /// <summary>
    ///     has column output formatter
    /// </summary>
    /// <typeparam name="TEntity">entity type</typeparam>
    /// <typeparam name="TProperty">property type</typeparam>
    /// <param name="configuration">property configuration</param>
    /// <param name="formatter">column output formatter</param>
    /// <returns>property configuration</returns>
    public static IPropertyConfiguration<TEntity, TProperty> HasColumnOutputFormatter<TEntity, TProperty>(
        this IPropertyConfiguration<TEntity, TProperty> configuration, Func<TProperty?, object?>? formatter)
    {
        if (formatter is null)
        {
            return configuration.HasOutputFormatter(null);
        }

        return configuration.HasOutputFormatter((_, prop) => formatter.Invoke(prop));
    }
}
