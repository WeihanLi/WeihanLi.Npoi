using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    internal class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// PropertyConfigurationDictionary
        /// </summary>
        public IDictionary<PropertyInfo, IPropertyConfiguration> PropertyConfigurationDictionary { get; internal set; }

        public ExcelSetting ExcelSetting { get; }

        internal IList<FreezeSetting> FreezeSettings { get; set; }

        internal FilterSetting FilterSetting { get; set; }

        internal IList<ISheetConfiguration> SheetConfigurations { get; set; }

        public ExcelConfiguration() : this(null)
        {
        }

        public ExcelConfiguration(ExcelSetting setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, IPropertyConfiguration>();
            ExcelSetting = setting ?? new ExcelSetting();
            SheetConfigurations = new List<ISheetConfiguration>(ExcelConstants.MaxSheetNum / 16)
            {
                new SheetConfiguration()
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

        public IExcelConfiguration HasFilter(int firstColumn) => HasFilter(firstColumn, null);

        public IExcelConfiguration HasFilter(int firstColumn, int? lastColumn)
        {
            FilterSetting = new FilterSetting(firstColumn, lastColumn);
            return this;
        }

        #region Property

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TEntity"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="IPropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
        {
            var pc = new PropertyConfiguration();

            var propertyInfo = GetPropertyInfo(propertyExpression);

            PropertyConfigurationDictionary[propertyInfo] = pc;

            return pc;
        }

        private static PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
        {
            if (propertyExpression.NodeType != ExpressionType.Lambda)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            var lambda = (LambdaExpression)propertyExpression;

            var memberExpression = ExtractMemberExpression(lambda.Body);
            if (memberExpression == null)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }
            return memberExpression.Member.DeclaringType.GetProperty(memberExpression.Member.Name);
        }

        private static MemberExpression ExtractMemberExpression(Expression expression)
        {
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                return (MemberExpression)expression;
            }

            if (expression.NodeType == ExpressionType.Convert)
            {
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            }

            return null;
        }

        #endregion Property
    }
}
