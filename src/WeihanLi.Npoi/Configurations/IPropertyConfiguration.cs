using System;

namespace WeihanLi.Npoi.Configurations
{
    /// <summary>
    /// PropertyConfiguration
    /// </summary>
    public interface IPropertyConfiguration
    {
    }

    public interface IPropertyConfiguration<TProperty> : IPropertyConfiguration
    {
        /// <summary>
        /// HasColumnIndex
        /// </summary>
        /// <param name="index">index</param>
        /// <returns></returns>
        IPropertyConfiguration<TProperty> HasColumnIndex(int index);

        /// <summary>
        /// HasColumnTitle
        /// </summary>
        /// <param name="title">title</param>
        /// <returns></returns>
        IPropertyConfiguration<TProperty> HasColumnTitle(string title);

        /// <summary>
        /// HasColumnFormatter
        /// </summary>
        /// <param name="formatter">formatter</param>
        /// <returns></returns>
        IPropertyConfiguration<TProperty> HasColumnFormatter(string formatter);

        /// <summary>
        /// Ignored
        /// </summary>
        /// <returns></returns>
        IPropertyConfiguration<TProperty> Ignored();

        /// <summary>
        /// HasColumnFormatter
        /// </summary>
        /// <param name="formatterFunc">columnFormatter</param>
        /// <returns></returns>
        IPropertyConfiguration<TProperty> HasColumnFormatter(Func<TProperty, object> formatterFunc);
    }
}
