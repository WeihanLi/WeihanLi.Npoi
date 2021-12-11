using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using OfficeOpenXml;

namespace WeihanLi.Npoi.Benchmark;

[SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5)]
[MemoryDiagnoser]
[MinColumn, MaxColumn, MeanColumn, MedianColumn]
public class WorkbookBasicTest
{
    private const int ColsCount = 10;

    [Params(10000, 30000, 50000, 65535)]
    public int RowsCount;

    [Benchmark(Baseline = true)]
    public byte[] NpoiXlsWorkbookInit()
    {
        var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xls);

        var sheet = workbook.CreateSheet("tempSheet");

        for (var i = 0; i < RowsCount; i++)
        {
            var row = sheet.CreateRow(i);
            for (var j = 0; j < ColsCount; j++)
            {
                var cell = row.CreateCell(j);
                cell.SetCellValue($"as ({i}, {j}) sa");
            }
        }

        return workbook.ToExcelBytes();
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoInlining)]
    public byte[] NpoiXlsxWorkbookInit()
    {
        var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);

        var sheet = workbook.CreateSheet("tempSheet");

        for (var i = 0; i < RowsCount; i++)
        {
            var row = sheet.CreateRow(i);
            for (var j = 0; j < ColsCount; j++)
            {
                var cell = row.CreateCell(j);
                cell.SetCellValue($"as ({i}, {j}) sa");
            }
        }

        return workbook.ToExcelBytes();
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoInlining)]
    public byte[] EpplusWorkbookInit()
    {
        var excel = new ExcelPackage();

        var sheet = excel.Workbook.Worksheets.Add("tempSheet");

        for (var i = 1; i <= RowsCount; i++)
        {
            for (var j = 1; j <= ColsCount; j++)
            {
                sheet.Cells[i, j].Value = $"as ({i}, {j}) sa";
            }
        }

        return excel.GetAsByteArray();
    }
}
