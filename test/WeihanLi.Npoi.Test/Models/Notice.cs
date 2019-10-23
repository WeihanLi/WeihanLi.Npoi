using System;

namespace WeihanLi.Npoi.Test.Models
{
    internal class Notice
    {
        public int Id { get; set; }

        public string Title { get; set; }

        public string Content { get; set; }

        public DateTime PublishedAt { get; set; }

        public string Publisher { get; set; }
    }
}
