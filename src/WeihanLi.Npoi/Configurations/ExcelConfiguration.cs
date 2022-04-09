// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System;
using System.Collections.Generic;
using System.Linq;
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

    public ExcelSetting ExcelSetting { get; } = ExcelHelper.DefaultExcelSetting;

    public IList<FreezeSetting> FreezeSettings { get; } = new List<FreezeSetting>();

    public FilterSetting? FilterSetting { get; set; }

    public IDictionary<int, SheetSetting> SheetSettings { get; } =
        new Dictionary<int, SheetSetting> { { 0, new SheetSetting() } };

    #region ExcelSettings FluentAPI

#nullable disable

    public IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction)
    {
        configAction?.Invoke(ExcelSetting);
        return this;
    }

#nullable restore

    #endregion ExcelSettings FluentAPI

    #region Sheet

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

    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit));
        return this;
    }

    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit, leftmostColumn, topRow));
        return this;
    }

    #endregion FreezePane

    #region Filter

    public IExcelConfiguration HasFilter(int firstColumn) => HasFilter(firstColumn, null);

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
    ///     EntityType
    /// </summary>
    public Type EntityType => typeof(TEntity);

    internal Func<TEntity?, bool>? DataFilter { get; private set; }

    internal IValidator? Validator { get; private set; }

    #region Property

    public IExcelConfiguration<TEntity> WithValidator(IValidator? validator)
    {
        Validator = validator;
        return this;
    }

    public IExcelConfiguration<TEntity> WithDataFilter(Func<TEntity?, bool>? dataFilter)
    {
        DataFilter = dataFilter;
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
        var property = memberInfo as PropertyInfo;
        if (property is null || !PropertyConfigurationDictionary.ContainsKey(property))
        {
            property = CacheUtil.GetTypeProperties(EntityType)
                .FirstOrDefault(p => p.Name == memberInfo.Name);
            if (null == property)
            {
                throw new InvalidOperationException($"the property [{memberInfo.Name}] does not exists");
            }
        }

        return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
    }

    public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName)
    {
        var property = PropertyConfigurationDictionary.Keys.FirstOrDefault(p => p.Name == propertyName);
        if (property != null)
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
