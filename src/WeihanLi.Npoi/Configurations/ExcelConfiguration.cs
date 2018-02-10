using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    public class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// PropertyConfigurationDictionary
        /// </summary>
        public IDictionary<PropertyInfo, IPropertyConfiguration> PropertyConfigurationDictionary { get; internal set; }

        public ExcelSetting ExcelSetting { get; }

        public IList<ISheetConfiguration> SheetConfigurations { get; internal set; }

        public ExcelConfiguration() : this(null)
        {
        }

        public ExcelConfiguration(ExcelSetting setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, IPropertyConfiguration>();
            ExcelSetting = setting ?? new ExcelSetting();
            SheetConfigurations = new List<ISheetConfiguration>(ExcelConstants.MaxSheetNum / 2)
            {
                new SheetConfiguration()
            };
        }

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

        private PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
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

        private MemberExpression ExtractMemberExpression(Expression expression)
        {
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                return ((MemberExpression)expression);
            }

            if (expression.NodeType == ExpressionType.Convert)
            {
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            }

            return null;
        }
    }
}
