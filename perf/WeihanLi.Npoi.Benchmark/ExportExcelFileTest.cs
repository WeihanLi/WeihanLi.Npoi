using System;
using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using EPPlus.Core.Extensions;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 3, targetCount: 10)]
    [MinColumn, MaxColumn, MeanColumn, MedianColumn]
    public class ExportExcelFileTest
    {
        private class TestEntity
        {
            public int PKID { get; set; }

            public string Username { get; set; }

            public string PasswordHash { get; set; }

            public decimal Amount { get; set; } = 1000M;

            public string WechatOpenId { get; set; }

            public bool IsActive { get; set; }

            public DateTime CreateTime { get; set; } = DateTime.Now;
        }

        private readonly string filePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\benchmarktest.xlsx";

        private readonly List<TestEntity> testData = new List<TestEntity>(102400);

        [GlobalSetup]
        public void GlobalSetup()
        {
            for (int i = 1; i <= 100_000; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i,
                });
            }
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            // Disposing logic
            testData.Clear();
        }

        [Benchmark]
        public void NpoiExportToFileTest()
        {
            testData.ToExcelFile(filePath.Replace(".xlsx", ".npoi.xlsx"));
        }

        [Benchmark]
        public void EpplusExportToFileTest()
        {
            testData.ToExcelPackage().SaveAs(new System.IO.FileInfo(filePath.Replace(".xlsx", ".epplus.xlsx")));
        }
    }
}
