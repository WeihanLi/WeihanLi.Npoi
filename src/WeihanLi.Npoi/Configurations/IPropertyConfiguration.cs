using System;

namespace WeihanLi.Npoi.Configurations
{
    /// <summary>
    /// PropertyConfiguration
    /// </summary>
    public interface IPropertyConfiguration
    {
    }

    public interface IPropertyConfiguration<TEntity, TProperty> : IPropertyConfiguration
    {
        /// <summary>
        /// HasColumnIndex
        /// </summary>
        /// <param name="index">index</param>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index);

        /// <summary>
        /// HasColumnWidth
        /// </summary>
        /// <param name="width">width</param>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width);

        /// <summary>
        /// HasColumnTitle
        /// </summary>
        /// <param name="title">title</param>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title);

        /// <summary>
        /// HasColumnFormatter
        /// </summary>
        /// <param name="formatter">formatter</param>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter);

        /// <summary>
        /// Ignored
        /// </summary>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> Ignored();

        /// <summary>
        /// HasColumnFormatter
        /// </summary>
        /// <param name="formatterFunc">columnFormatter</param>
        /// <returns></returns>
        IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(Func<TEntity, TProperty, object> formatterFunc);
    }
}
