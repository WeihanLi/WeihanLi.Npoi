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
        [Benchmark]
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void NpoiXlsWorkbookInit()
        {
            var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xls);

            var sheet = workbook.CreateSheet("tempSheet");

            for (int i = 0; i < 50000; i++)
            {
                var row = sheet.CreateRow(i);
                for (int j = 0; j < 10; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.SetCellValue("asasasassa");
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

            for (int i = 0; i < 50000; i++)
            {
                var row = sheet.CreateRow(i);
                for (int j = 0; j < 10; j++)
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

            for (int i = 1; i <= 50000; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    sheet.Cells[i, j].Value = $"as ({i}, {j}) sa";
                }
            }

            var excelBytes = excel.GetAsByteArray();
        }
    }
}
