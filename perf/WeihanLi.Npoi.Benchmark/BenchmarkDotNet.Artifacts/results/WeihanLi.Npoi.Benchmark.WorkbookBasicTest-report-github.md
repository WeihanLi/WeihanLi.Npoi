``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-DEIOEH : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|               Method |    Mean |    Error |   StdDev |     Min |     Max |  Median |       Gen 0 |      Gen 1 |     Gen 2 |  Allocated |
|--------------------- |--------:|---------:|---------:|--------:|--------:|--------:|------------:|-----------:|----------:|-----------:|
|  NpoiXlsWorkbookInit | 3.731 s | 0.0485 s | 0.0321 s | 3.700 s | 3.799 s | 3.722 s |  62000.0000 | 22000.0000 | 5000.0000 |  504.45 MB |
| NpoiXlsxWorkbookInit | 9.257 s | 0.1235 s | 0.0817 s | 9.094 s | 9.364 s | 9.285 s | 386000.0000 | 67000.0000 | 7000.0000 | 2048.14 MB |
|   EpplusWorkbookInit | 3.699 s | 0.0208 s | 0.0137 s | 3.669 s | 3.714 s | 3.701 s | 100000.0000 | 32000.0000 | 7000.0000 |  747.84 MB |
