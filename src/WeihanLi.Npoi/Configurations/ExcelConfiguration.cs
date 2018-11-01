using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// EntityType
        /// </summary>
        private readonly Type _entityType = typeof(TEntity);

        /// <summary>
        /// PropertyConfigurationDictionary
        /// </summary>
        public IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigurationDictionary { get; internal set; }

        public ExcelSetting ExcelSetting { get; }

        internal IList<FreezeSetting> FreezeSettings { get; set; }

        internal FilterSetting FilterSetting { get; set; }

        internal IList<SheetSetting> SheetSettings { get; set; }

        public ExcelConfiguration() : this(null)
        {
        }

        public ExcelConfiguration(ExcelSetting setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, PropertyConfiguration>();
            ExcelSetting = setting ?? new ExcelSetting();
            SheetSettings = new List<SheetSetting>(InternalConstants.MaxSheetNum / 16)
            {
                new SheetSetting()
            };
            FreezeSettings = new List<FreezeSetting>();
        }

        #region ExcelSettings FluentAPI

        public IExcelConfiguration HasAuthor(string author)
        {
            ExcelSetting.Author = author;
            return this;
        }

        public IExcelConfiguration HasTitle(string title)
        {
            ExcelSetting.Title = title;
            return this;
        }

        public IExcelConfiguration HasDescription(string description)
        {
            ExcelSetting.Description = description;
            return this;
        }

        public IExcelConfiguration HasSubject(string subject)
        {
            ExcelSetting.Subject = subject;
            return this;
        }

        public IExcelConfiguration HasCompany(string company)
        {
            ExcelSetting.Company = company;
            return this;
        }

        public IExcelConfiguration HasCategory(string category)
        {
            ExcelSetting.Category = category;
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

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TEntity"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="IPropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression) => PropertyConfigurationDictionary[_entityType.GetProperty(propertyExpression.GetMemberInfo().Name)];

        #endregion Property

        #region Sheet

        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName) =>
            HasSheetConfiguration(sheetIndex, sheetName, 1);

        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex)
        {
            var sheetSetting =
                SheetSettings.FirstOrDefault(_ => _.SheetIndex == sheetIndex);

            if (sheetSetting == null)
            {
                SheetSettings.Add(new SheetSetting
                {
                    SheetIndex = sheetIndex,
                    SheetName = sheetName,
                    StartRowIndex = startRowIndex
                });
            }
            else
            {
                sheetSetting.SheetName = sheetName;
                sheetSetting.StartRowIndex = startRowIndex;
            }

            return this;
        }

        #endregion Sheet
    }
}
