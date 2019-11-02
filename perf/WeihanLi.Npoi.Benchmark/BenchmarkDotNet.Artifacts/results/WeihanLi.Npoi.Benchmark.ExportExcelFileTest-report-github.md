``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-THIRET : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                        Method |    Mean |    Error |   StdDev |     Min |     Max |  Median |      Gen 0 |      Gen 1 |     Gen 2 | Allocated |
|------------------------------ |--------:|---------:|---------:|--------:|--------:|--------:|-----------:|-----------:|----------:|----------:|
|         NpoiExportToBytesTest | 2.562 s | 0.0191 s | 0.0126 s | 2.537 s | 2.580 s | 2.561 s | 97000.0000 | 23000.0000 | 5000.0000 | 534.44 MB |
|      NpoiExportToXlsBytesTest | 1.702 s | 0.0174 s | 0.0115 s | 1.687 s | 1.725 s | 1.703 s | 34000.0000 | 12000.0000 | 3000.0000 | 214.15 MB |
|       EpplusExportToBytesTest | 1.043 s | 0.0077 s | 0.0051 s | 1.032 s | 1.048 s | 1.043 s | 51000.0000 | 13000.0000 | 2000.0000 | 270.93 MB |
|   NpoiStructExportToBytesTest | 2.573 s | 0.0346 s | 0.0229 s | 2.542 s | 2.606 s | 2.572 s | 97000.0000 | 23000.0000 | 5000.0000 | 536.35 MB |
| EpplusStructExportToBytesTest | 1.060 s | 0.0086 s | 0.0051 s | 1.052 s | 1.069 s | 1.058 s | 51000.0000 | 13000.0000 | 2000.0000 | 270.94 MB |
