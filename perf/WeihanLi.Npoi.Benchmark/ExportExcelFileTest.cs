using BenchmarkDotNet.Attributes;
using EPPlus.Core.Extensions;
using EPPlus.Core.Extensions.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 3, targetCount: 10)]
    [MinColumn, MaxColumn, MeanColumn, MedianColumn]
    public class ExportExcelFileTest
    {
        private class TestEntity
        {
            [ExcelTableColumn("PKID")]
            public int PKID { get; set; }

            [ExcelTableColumn("UserName")]
            public string Username { get; set; }

            [ExcelTableColumn("PasswordHash")]
            public string PasswordHash { get; set; }

            [ExcelTableColumn("Amount")]
            public decimal Amount { get; set; } = 1000M;

            [ExcelTableColumn("WechatOpenId")]
            public string WechatOpenId { get; set; }

            [ExcelTableColumn("IsActive")]
            public bool IsActive { get; set; }

            [ExcelTableColumn("CreateTime")]
            public DateTime CreateTime { get; set; } = DateTime.Now;
        }

        private readonly string filePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\benchmarktest.xlsx";

        private readonly List<TestEntity> testData = new List<TestEntity>(102400);

        [GlobalSetup]
        public void GlobalSetup()
        {
            for (var i = 1; i <= 100_000; i++)
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
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiExportToFileTest()
        {
            testData.ToExcelFile(filePath.Replace(".xlsx", ".npoi.xlsx"));
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusExportToFileTest()
        {
            testData.ToExcelPackage().SaveAs(new System.IO.FileInfo(filePath.Replace(".xlsx", ".epplus.xlsx")));
        }
    }
}
