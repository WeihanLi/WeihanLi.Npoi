// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WeihanLi.Npoi.Benchmark;

[SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5)]
[MemoryDiagnoser]
[MinColumn, MaxColumn, MeanColumn, MedianColumn]
public class ImportExcelTest
{
    private class TestEntity
    {
        public int PKID { get; set; }

        public string? Username { get; set; }

        public string? PasswordHash { get; set; }

        public decimal Amount { get; set; }

        public string? WechatOpenId { get; set; }

        public bool IsActive { get; set; }

        public DateTime CreateTime { get; set; } = DateTime.Now;
    }

    private readonly List<TestEntity> testData = new(51200);
    private byte[] xlsBytes = Array.Empty<byte>(), xlsxBytes = Array.Empty<byte>(), csvBytes = Array.Empty<byte>();

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
        }

        xlsBytes = testData.ToExcelBytes();
        xlsxBytes = testData.ToExcelBytes(ExcelFormat.Xlsx);
        csvBytes = testData.ToCsvBytes();
    }

    [GlobalCleanup]
    public void GlobalCleanup()
    {
        // Disposing logic
        testData.Clear();
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
