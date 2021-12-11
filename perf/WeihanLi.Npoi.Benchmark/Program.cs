using BenchmarkDotNet.Running;

namespace WeihanLi.Npoi.Benchmark;

public class Program
{
    public static void Main(string[] args)
    {
        // BenchmarkRunner.Run<WorkbookBasicTest>();
        BenchmarkRunner.Run<ExportExcelTest>();
        BenchmarkRunner.Run<ImportExcelTest>();
    }
}
