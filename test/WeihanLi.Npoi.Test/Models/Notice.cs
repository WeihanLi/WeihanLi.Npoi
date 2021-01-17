using System;

namespace WeihanLi.Npoi.Test.Models
{
    internal class BaseModel
    {
        public int Id { get; set; }
    }

    internal class Notice : BaseModel
    {
        public string? Title { get; set; }

        public string? Content { get; set; }

        public DateTime PublishedAt { get; set; }

        public string? Publisher { get; set; }
    }
}
