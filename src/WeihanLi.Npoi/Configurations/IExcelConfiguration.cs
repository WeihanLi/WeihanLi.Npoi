using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    public interface IExcelConfiguration
    {
        IDictionary<PropertyInfo, IPropertyConfiguration> PropertyConfigurationDictionary { get; }

        ExcelSetting ExcelSetting { get; }

        IList<ISheetConfiguration> SheetConfigurations { get; }
    }

    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        #region ExcelSettings FluentAPI

        IExcelConfiguration<TEntity> HasAuthor(string author);

        IExcelConfiguration<TEntity> HasTitle(string title);

        IExcelConfiguration<TEntity> HasDescription(string description);

        IExcelConfiguration<TEntity> HasSubject(string subject);

        #endregion ExcelSettings FluentAPI
    }
}
