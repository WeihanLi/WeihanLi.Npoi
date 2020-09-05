using BenchmarkDotNet.Attributes;
using EPPlus.Core.Extensions;
using EPPlus.Core.Extensions.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using WeihanLi.Document;
using WeihanLi.Document.Excel;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5)]
    [MemoryDiagnoser]
    [MinColumn, MaxColumn, MeanColumn, MedianColumn]
    public class ExportExcelTest
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

        private readonly List<TestEntity> testData = new List<TestEntity>(51200);
        private readonly List<TestStruct> testStructData = new List<TestStruct>(51200);

        [Params(10000, 30000, 50000, 65535)]
        public int RowsCount;

        [GlobalSetup]
        public void GlobalSetup()
        {
            for (var i = 1; i <= RowsCount; i++)
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

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void ExportToCsvBytesTest()
        {
            var bytes = testData.ToCsvBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiExportToXlsBytesTest()
        {
            var excelBytes = testData.ToExcelBytes(ExcelFormat.Xls);
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiExportToXlsxBytesTest()
        {
            var excelBytes = testData.ToExcelBytes(ExcelFormat.Xlsx);
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusExportToBytesTest()
        {
            var excelBytes = testData.ToExcelPackage().GetAsByteArray();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void StructExportToCsvBytesTest()
        {
            var bytes = testStructData.ToCsvBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiStructExportToXlsBytesTest()
        {
            var excelBytes = testStructData.ToExcelBytes(ExcelFormat.Xls);
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiStructExportToXlsxBytesTest()
        {
            var excelBytes = testStructData.ToExcelBytes(ExcelFormat.Xlsx);
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusStructExportToBytesTest()
        {
            var excelBytes = testStructData.ToExcelPackage().GetAsByteArray();
        }
    }
}
