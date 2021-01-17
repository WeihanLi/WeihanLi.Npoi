using System;

// ReSharper disable InconsistentNaming
namespace DotNetCoreSample
{
    internal class ProductPriceMapping
    {
        public string? Type { get; set; }

        public string? Pid { get; set; }

        public string? ShopCode { get; set; }

        public decimal Price { get; set; }

        public long ItemID { get; set; }

        public long SkuID { get; set; }

        public DateTime LastUpdateDateTime { get; set; }
    }
}
