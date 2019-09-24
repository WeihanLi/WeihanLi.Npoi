using BenchmarkDotNet.Running;

namespace WeihanLi.Npoi.Benchmark
{
    public class Program
    {
        public static void Main(string[] args)
        {
            BenchmarkRunner.Run<ExportExcelFileTest>();
        }
    }
}
