using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal sealed class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// EntityType
        /// </summary>
        public Type EntityType => typeof(TEntity);

        /// <summary>
        /// PropertyConfigurationDictionary
        /// </summary>
        public IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigurationDictionary { get; internal set; }

        public ExcelSetting ExcelSetting { get; }

        internal IList<FreezeSetting> FreezeSettings { get; set; }

        internal FilterSetting? FilterSetting { get; set; }

        internal IDictionary<int, SheetSetting> SheetSettings { get; set; }

        internal Func<TEntity?, bool>? DataValidationFunc { get; private set; }

        public ExcelConfiguration() : this(null)
        {
        }

        public ExcelConfiguration(ExcelSetting? setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, PropertyConfiguration>();
            ExcelSetting = setting ?? ExcelHelper.DefaultExcelSetting;
            SheetSettings = new Dictionary<int, SheetSetting>(4)
            {
                { 0, new SheetSetting() }
            };
            FreezeSettings = new List<FreezeSetting>(4);
        }

        #region ExcelSettings FluentAPI

        public IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction)
        {
            configAction.Invoke(ExcelSetting);
            return this;
        }

        #endregion ExcelSettings FluentAPI

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

        #region Property

        public IExcelConfiguration<TEntity> WithDataValidation(Func<TEntity?, bool>? dataValidateFunc)
        {
            DataValidationFunc = dataValidateFunc;
            return this;
        }

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TEntity"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="IPropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
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
            var propertyConfiguration = (PropertyConfiguration)Activator.CreateInstance(propertyConfigurationType, new object[] { property });

            PropertyConfigurationDictionary[property] = propertyConfiguration;

            return (IPropertyConfiguration<TEntity, TProperty>)propertyConfiguration;
        }

        #endregion Property

        #region Sheet

        public IExcelConfiguration HasSheetSetting(Action<SheetSetting> configAction, int sheetIndex = 0)
        {
            if (sheetIndex >= 0)
            {
                if (SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
                {
                    configAction.Invoke(sheetSetting);
                }
                else
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
    }
}
