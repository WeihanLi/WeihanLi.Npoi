``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-ABUDTI : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                 Method |    Mean |    Error |   StdDev |     Min |     Max |  Median |
|----------------------- |--------:|---------:|---------:|--------:|--------:|--------:|
|   NpoiExportToFileTest | 5.220 s | 0.4123 s | 0.2727 s | 4.952 s | 5.596 s | 5.162 s |
| EpplusExportToFileTest | 2.023 s | 0.0138 s | 0.0091 s | 2.009 s | 2.036 s | 2.025 s |
