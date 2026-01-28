// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Common.Services;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations;

internal abstract class ExcelConfiguration : IExcelConfiguration
{
    /// <summary>
    ///     PropertyConfigurationDictionary
    /// </summary>
    public IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigurationDictionary { get; } =
        new Dictionary<PropertyInfo, PropertyConfiguration>();

    /// <summary>
    ///     Gets the Excel-level document metadata.
    /// </summary>
    public ExcelSetting ExcelSetting { get; } = ExcelHelper.DefaultExcelSetting;

    /// <summary>
    ///     Gets the configured freeze panes for the workbook.
    /// </summary>
    public IList<FreezeSetting> FreezeSettings { get; } = new List<FreezeSetting>();

    /// <summary>
    ///     Gets or sets the filter configuration for the sheet.
    /// </summary>
    public FilterSetting? FilterSetting { get; set; }

    /// <summary>
    ///     Gets the registered sheet settings keyed by sheet index.
    /// </summary>
    public IDictionary<int, SheetSetting> SheetSettings { get; } =
        new Dictionary<int, SheetSetting> { { 0, new SheetSetting() } };

    #region ExcelSettings FluentAPI



    /// <summary>
    ///     Updates the Excel document metadata in a fluent manner.
    /// </summary>
    /// <param name="configAction">Configuration delegate.</param>
    /// <returns>The current configuration instance.</returns>
    public IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction)
    {
        // ReSharper disable once ConditionalAccessQualifierIsNonNullableAccordingToAPIContract
        // allow nullable but we do not want a null
        configAction?.Invoke(ExcelSetting);
        return this;
    }

    #endregion ExcelSettings FluentAPI

    #region Sheet

    /// <summary>
    ///     Updates the sheet configuration for the specified index.
    /// </summary>
    /// <param name="configAction">Configuration delegate.</param>
    /// <param name="sheetIndex">Target sheet index.</param>
    /// <returns>The current configuration instance.</returns>
    public IExcelConfiguration HasSheetSetting(Action<SheetSetting> configAction, int sheetIndex = 0)
    {
        if (configAction is null)
        {
            throw new ArgumentNullException(nameof(configAction));
        }

        if (sheetIndex >= 0)
        {
            if (!SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
            {
                SheetSettings[sheetIndex]
                    = sheetSetting
                        = new SheetSetting();
            }

            configAction.Invoke(sheetSetting);
        }

        return this;
    }

    #endregion Sheet

    #region FreezePane

    /// <summary>
    ///     Adds a freeze pane using default anchor values.
    /// </summary>
    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit));
        return this;
    }

    /// <summary>
    ///     Adds a freeze pane with the specified anchor.
    /// </summary>
    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit, leftmostColumn, topRow));
        return this;
    }

    #endregion FreezePane

    #region Filter

    /// <summary>
    ///     Adds an auto-filter that starts at the specified column.
    /// </summary>
    public IExcelConfiguration HasFilter(int firstColumn) => HasFilter(firstColumn, null);

    /// <summary>
    ///     Adds an auto-filter covering the specified column range.
    /// </summary>
    public IExcelConfiguration HasFilter(int firstColumn, int? lastColumn)
    {
        FilterSetting = new FilterSetting(firstColumn, lastColumn);
        return this;
    }

    #endregion Filter
}

internal sealed class ExcelConfiguration<TEntity> : ExcelConfiguration, IExcelConfiguration<TEntity>
{
    /// <summary>
    ///     Gets the entity type represented by this configuration.
    /// </summary>
    public Type EntityType => typeof(TEntity);

    internal Func<TEntity?, bool>? DataFilter { get; private set; }

    internal Action<TEntity?, int>? PostImportAction { get; private set; }

    internal IComparer<PropertyInfo>? PropertyComparer { get; private set; }

    internal IValidator? Validator { get; private set; }

    #region Property

    /// <summary>
    ///     Assigns a custom validator.
    /// </summary>
    public IExcelConfiguration<TEntity> WithValidator(IValidator? validator)
    {
        Validator = validator;
        return this;
    }

    /// <summary>
    ///     Applies a data filter that can skip entities during export.
    /// </summary>
    public IExcelConfiguration<TEntity> WithDataFilter(Func<TEntity?, bool>? dataFilter)
    {
        DataFilter = dataFilter;
        return this;
    }

    public IExcelConfiguration<TEntity> WithPostImportAction(Action<TEntity?, int>? postAction)
    {
        PostImportAction = postAction;
        return this;
    }

    /// <summary>
    ///     Controls the ordering of properties.
    /// </summary>
    public IExcelConfiguration<TEntity> WithPropertyComparer(IComparer<PropertyInfo>? propertyComparer)
    {
        PropertyComparer = propertyComparer;
        return this;
    }

    /// <summary>
    ///     Gets the property configuration by the specified property expression for the specified
    ///     <typeparamref name="TEntity" /> and its <typeparamref name="TProperty" />.
    /// </summary>
    /// <returns>The <see cref="IPropertyConfiguration" />.</returns>
    /// <param name="propertyExpression">The property expression.</param>
    /// <typeparam name="TProperty">The type of parameter.</typeparam>
    public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(
        Expression<Func<TEntity, TProperty>> propertyExpression)
    {
        var memberInfo = propertyExpression.GetMemberInfo();
        if (memberInfo is PropertyInfo property &&
            PropertyConfigurationDictionary.TryGetValue(property, out var propertyConfiguration))
            return (IPropertyConfiguration<TEntity, TProperty>)propertyConfiguration;

        property = CacheUtil.GetTypeProperties(EntityType).FirstOrDefault(p => p.Name == memberInfo.Name)
                   ?? throw new InvalidOperationException($"the property [{memberInfo.Name}] does not exists");
        return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
    }

    /// <summary>
    ///     Retrieves (or builds) a configuration for the specified property name.
    /// </summary>
    /// <typeparam name="TProperty">Property type.</typeparam>
    /// <param name="propertyName">Property name.</param>
    /// <returns>Property configuration.</returns>
    public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName)
    {
        var property = PropertyConfigurationDictionary.Keys.FirstOrDefault(p => p.Name == propertyName);
        if (property is not null)
        {
            return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
        }

        var propertyType = typeof(TProperty);

        property = new FakePropertyInfo(EntityType, propertyType, propertyName);

        var propertyConfigurationType =
            typeof(PropertyConfiguration<,>).MakeGenericType(EntityType, propertyType);
        var propertyConfiguration =
            (PropertyConfiguration)Guard.NotNull(Activator.CreateInstance(propertyConfigurationType, property));

        PropertyConfigurationDictionary[property] = propertyConfiguration;

        return (IPropertyConfiguration<TEntity, TProperty>)propertyConfiguration;
    }

    #endregion Property
}
