``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-6300U CPU 2.40GHz (Skylake), 1 CPU, 4 logical and 2 physical cores
.NET Core SDK=2.2.300
  [Host]     : .NET Core 2.2.5 (CoreCLR 4.6.27617.05, CoreFX 4.6.27618.01), 64bit RyuJIT
  Job-EGKCKL : .NET Core 2.2.5 (CoreCLR 4.6.27617.05, CoreFX 4.6.27618.01), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                 Method |    Mean |   Error |  StdDev |     Min |     Max |  Median |
|----------------------- |--------:|--------:|--------:|--------:|--------:|--------:|
|   NpoiExportToFileTest | 27.29 s | 1.881 s | 1.244 s | 24.63 s | 28.68 s | 27.47 s |
| EpplusExportToFileTest |      NA |      NA |      NA |      NA |      NA |      NA |

Benchmarks with issues:
  ExportExcelFileTest.EpplusExportToFileTest: Job-EGKCKL(IterationCount=10, LaunchCount=1, WarmupCount=3)
