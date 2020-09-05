using System;
using System.Linq.Expressions;

namespace WeihanLi.Document.Configurations
{
    public interface IDocumentConfiguration
    {
    }

    public interface IDocumentConfiguration<TEntity> : IDocumentConfiguration
    {
        /// <summary>
        /// property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyExpression">propertyExpression to get property info</param>
        /// <returns>current excel configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        /// property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyName">propertyName</param>
        /// <returns>current excel configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
