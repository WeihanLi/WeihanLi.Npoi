namespace WeihanLi.Npoi.Configurations
{
    /// <summary>
    /// PropertyConfiguration
    /// </summary>
    public interface IPropertyConfiguration
    {
        /// <summary>
        /// HasColumnIndex
        /// </summary>
        /// <param name="index">index</param>
        /// <returns></returns>
        IPropertyConfiguration HasColumnIndex(int index);

        /// <summary>
        /// HasColumnTitle
        /// </summary>
        /// <param name="title">title</param>
        /// <returns></returns>
        IPropertyConfiguration HasColumnTitle(string title);

        /// <summary>
        /// HasColumnFormatter
        /// </summary>
        /// <param name="formatter">formatter</param>
        /// <returns></returns>
        IPropertyConfiguration HasColumnFormatter(string formatter);

        /// <summary>
        /// Ignored
        /// </summary>
        /// <returns></returns>
        IPropertyConfiguration Ignored();
    }
}
