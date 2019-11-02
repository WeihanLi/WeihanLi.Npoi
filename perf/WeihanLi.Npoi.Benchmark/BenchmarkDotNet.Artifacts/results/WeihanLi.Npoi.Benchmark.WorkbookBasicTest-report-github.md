``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-ALKOBF : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|               Method |    Mean |    Error |   StdDev |     Min |     Max |  Median |       Gen 0 |      Gen 1 |     Gen 2 |  Allocated |
|--------------------- |--------:|---------:|---------:|--------:|--------:|--------:|------------:|-----------:|----------:|-----------:|
|  NpoiXlsWorkbookInit | 1.652 s | 0.0036 s | 0.0021 s | 1.650 s | 1.657 s | 1.652 s |  36000.0000 | 13000.0000 | 4000.0000 |  234.95 MB |
| NpoiXlsxWorkbookInit | 7.183 s | 0.1493 s | 0.0988 s | 7.090 s | 7.329 s | 7.146 s | 287000.0000 | 51000.0000 | 6000.0000 | 1545.32 MB |
|   EpplusWorkbookInit | 2.866 s | 0.0260 s | 0.0155 s | 2.847 s | 2.893 s | 2.863 s |  78000.0000 | 25000.0000 | 7000.0000 |  578.46 MB |
