using BenchmarkDotNet.Attributes;
using OfficeOpenXml;
using System.Runtime.CompilerServices;

namespace WeihanLi.Npoi.Benchmark
{
    [SimpleJob(launchCount: 1, warmupCount: 3, targetCount: 10)]
    [MemoryDiagnoser]
    [MinColumn, MaxColumn, MeanColumn, MedianColumn]
    public class WorkbookBasicTest
    {
        private const int RowsCount = 65535;
        private const int ColsCount = 10;

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiXlsWorkbookInit()
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

            var excelBytes = workbook.ToExcelBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiXlsxWorkbookInit()
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

            var excelBytes = workbook.ToExcelBytes();
        }

        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void EpplusWorkbookInit()
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

            var excelBytes = excel.GetAsByteArray();
        }
    }
}
