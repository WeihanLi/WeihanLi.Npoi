``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-VOGULL : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                  Method |    Mean |    Error |   StdDev |     Min |     Max |  Median |      Gen 0 |      Gen 1 |     Gen 2 | Allocated |
|------------------------ |--------:|---------:|---------:|--------:|--------:|--------:|-----------:|-----------:|----------:|----------:|
|  ImportFromXlsBytesTest | 1.659 s | 0.0186 s | 0.0123 s | 1.637 s | 1.674 s | 1.663 s | 52000.0000 | 16000.0000 | 2000.0000 | 253.25 MB |
| ImportFromXlsxBytesTest | 3.786 s | 0.0482 s | 0.0319 s | 3.745 s | 3.853 s | 3.779 s | 93000.0000 | 35000.0000 | 3000.0000 | 490.18 MB |
