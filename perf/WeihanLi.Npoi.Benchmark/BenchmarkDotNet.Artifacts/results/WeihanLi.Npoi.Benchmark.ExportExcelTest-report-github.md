``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-RTOLTK : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=10  LaunchCount=1  WarmupCount=3  

```
|                          Method |       Mean |      Error |     StdDev |        Min |        Max |     Median |       Gen 0 |      Gen 1 |     Gen 2 | Allocated |
|-------------------------------- |-----------:|-----------:|-----------:|-----------:|-----------:|-----------:|------------:|-----------:|----------:|----------:|
|            ExportToCsvBytesTest |   121.2 ms |  2.2688 ms |  1.5007 ms |   119.8 ms |   124.8 ms |   120.9 ms |  10000.0000 |  2000.0000 |  800.0000 |  60.81 MB |
|        NpoiExportToXlsBytesTest | 1,686.8 ms | 15.1307 ms | 10.0080 ms | 1,674.0 ms | 1,700.0 ms | 1,685.6 ms |  33000.0000 | 12000.0000 | 3000.0000 | 212.25 MB |
|       NpoiExportToXlsxBytesTest | 2,483.3 ms | 14.0983 ms |  8.3897 ms | 2,469.6 ms | 2,492.9 ms | 2,487.8 ms |  95000.0000 | 22000.0000 | 4000.0000 | 532.54 MB |
|         EpplusExportToBytesTest | 1,027.0 ms |  3.9548 ms |  2.6159 ms | 1,022.8 ms | 1,029.7 ms | 1,027.4 ms |  51000.0000 | 13000.0000 | 2000.0000 | 270.94 MB |
|      StructExportToCsvBytesTest |   113.2 ms |  0.5914 ms |  0.3912 ms |   112.7 ms |   113.9 ms |   113.2 ms |  10000.0000 |  2000.0000 |  800.0000 |  60.81 MB |
|  NpoiStructExportToXlsBytesTest | 1,653.8 ms | 10.5413 ms |  6.9724 ms | 1,644.3 ms | 1,665.4 ms | 1,654.6 ms |  33000.0000 | 12000.0000 | 3000.0000 | 214.15 MB |
| NpoiStructExportToXlsxBytesTest | 2,513.1 ms | 20.8778 ms | 13.8094 ms | 2,482.7 ms | 2,525.5 ms | 2,518.7 ms | 100000.0000 | 24000.0000 | 4000.0000 | 534.44 MB |
|   EpplusStructExportToBytesTest | 1,045.9 ms |  8.0583 ms |  5.3301 ms | 1,037.8 ms | 1,054.1 ms | 1,046.7 ms |  51000.0000 | 13000.0000 | 2000.0000 | 270.95 MB |
