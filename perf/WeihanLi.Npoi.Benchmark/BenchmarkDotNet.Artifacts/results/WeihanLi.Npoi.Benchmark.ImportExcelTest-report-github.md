``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-YRZGQS : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                  Method |       Mean |    Error |   StdDev |        Min |        Max |     Median |       Gen 0 |      Gen 1 |     Gen 2 | Allocated |
|------------------------ |-----------:|---------:|---------:|-----------:|-----------:|-----------:|------------:|-----------:|----------:|----------:|
|  ImportFromCsvBytesTest |   894.6 ms | 17.91 ms | 11.85 ms |   881.2 ms |   918.5 ms |   897.7 ms |  50000.0000 | 12000.0000 | 1000.0000 |    182 MB |
|  ImportFromXlsBytesTest | 1,722.5 ms | 29.35 ms | 19.42 ms | 1,694.0 ms | 1,753.2 ms | 1,719.0 ms |  76000.0000 | 28000.0000 | 2000.0000 | 315.85 MB |
| ImportFromXlsxBytesTest | 3,997.9 ms | 61.59 ms | 40.74 ms | 3,952.2 ms | 4,062.1 ms | 3,986.1 ms | 110000.0000 | 36000.0000 | 3000.0000 | 552.76 MB |
