using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 3, targetCount: 10)]
    [MemoryDiagnoser]
    [MinColumn, MaxColumn, MeanColumn, MedianColumn]
    public class ImportExcelTest
    {
        private class TestEntity
        {
            public int PKID { get; set; }

            public string Username { get; set; }

            public string PasswordHash { get; set; }

            public decimal Amount { get; set; }

            public string WechatOpenId { get; set; }

            public bool IsActive { get; set; }

            public DateTime CreateTime { get; set; } = DateTime.Now;
        }

        private readonly List<TestEntity> testData = new List<TestEntity>(51200);
        private byte[] xlsBytes, xlsxBytes, csvBytes;

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
            }

            xlsBytes = testData.ToExcelBytes(ExcelFormat.Xls);
            xlsxBytes = testData.ToExcelBytes(ExcelFormat.Xlsx);
            csvBytes = testData.ToCsvBytes();
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            // Disposing logic
            testData.Clear();
            xlsBytes = null;
            xlsxBytes = null;
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public int ImportFromCsvBytesTest()
        {
            var list = CsvHelper.ToEntityList<TestEntity>(csvBytes);
            return list.Count;
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public int ImportFromXlsBytesTest()
        {
            var list = ExcelHelper.ToEntityList<TestEntity>(xlsBytes, ExcelFormat.Xls);
            return list.Count;
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public int ImportFromXlsxBytesTest()
        {
            var list = ExcelHelper.ToEntityList<TestEntity>(xlsxBytes, ExcelFormat.Xlsx);
            return list.Count;
        }
    }
}
