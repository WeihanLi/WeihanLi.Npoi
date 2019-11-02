using BenchmarkDotNet.Attributes;
using EPPlus.Core.Extensions;
using EPPlus.Core.Extensions.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 3, targetCount: 10)]
    [MemoryDiagnoser]
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
            public decimal Amount { get; set; }

            [ExcelTableColumn("WechatOpenId")]
            public string WechatOpenId { get; set; }

            [ExcelTableColumn("IsActive")]
            public bool IsActive { get; set; }

            [ExcelTableColumn("CreateTime")]
            public DateTime CreateTime => DateTime.Now;
        }

        private struct TestStruct
        {
            [ExcelTableColumn("PKID")]
            public int PKID { get; set; }

            [ExcelTableColumn("UserName")]
            public string Username { get; set; }

            [ExcelTableColumn("PasswordHash")]
            public string PasswordHash { get; set; }

            [ExcelTableColumn("Amount")]
            public decimal Amount { get; set; }

            [ExcelTableColumn("WechatOpenId")]
            public string WechatOpenId { get; set; }

            [ExcelTableColumn("IsActive")]
            public bool IsActive { get; set; }

            [ExcelTableColumn("CreateTime")]
            public DateTime CreateTime => DateTime.Now;
        }

        private readonly string filePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\benchmarktest.xlsx";

        private readonly List<TestEntity> testData = new List<TestEntity>(51200);
        private readonly List<TestStruct> testStructData = new List<TestStruct>(51200);

        [GlobalSetup]
        public void GlobalSetup()
        {
            for (var i = 1; i <= 50_000; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    PKID = i,
                });

                testStructData.Add(new TestStruct()
                {
                    Amount = 1000,
                    Username = "xxxx",
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

        //[Benchmark]
        //[MethodImpl(MethodImplOptions.NoInlining)]
        //public void NpoiExportToFileTest()
        //{
        //    testData.ToExcelFile(filePath.Replace(".xlsx", ".npoi.xlsx"));
        //}

        //[Benchmark]
        //[MethodImpl(MethodImplOptions.NoInlining)]
        //public void EpplusExportToFileTest()
        //{
        //    testData.ToExcelPackage().SaveAs(new System.IO.FileInfo(filePath.Replace(".xlsx", ".epplus.xlsx")));
        //}

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiExportToBytesTest()
        {
            var excelBytes = testData.ToExcelBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiExportToXlsBytesTest()
        {
            var excelBytes = testData.ToExcelBytes(ExcelFormat.Xls);
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusExportToBytesTest()
        {
            var excelBytes = testData.ToExcelPackage().GetAsByteArray();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiStructExportToBytesTest()
        {
            var excelBytes = testStructData.ToExcelBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusStructExportToBytesTest()
        {
            var excelBytes = testStructData.ToExcelPackage().GetAsByteArray();
        }
    }
}
