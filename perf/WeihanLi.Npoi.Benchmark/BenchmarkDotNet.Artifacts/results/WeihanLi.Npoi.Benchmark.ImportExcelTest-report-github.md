``` ini

BenchmarkDotNet=v0.11.5, OS=Windows 10.0.18362
Intel Core i5-3470 CPU 3.20GHz (Ivy Bridge), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=3.0.100
  [Host]     : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT
  Job-WDPKYY : .NET Core 2.2.6 (CoreCLR 4.6.27817.03, CoreFX 4.6.27818.02), 64bit RyuJIT

IterationCount=5  LaunchCount=1  WarmupCount=1  

```
|                  Method | RowsCount |       Mean |      Error |     StdDev |        Min |        Max |     Median |       Gen 0 |      Gen 1 |     Gen 2 | Allocated |
|------------------------ |---------- |-----------:|-----------:|-----------:|-----------:|-----------:|-----------:|------------:|-----------:|----------:|----------:|
|  **ImportFromCsvBytesTest** |     **10000** |   **171.2 ms** |   **2.904 ms** |  **0.7541 ms** |   **170.0 ms** |   **172.0 ms** |   **171.2 ms** |   **9333.3333** |  **2000.0000** |  **333.3333** |  **36.44 MB** |
|  ImportFromXlsBytesTest |     10000 |   353.8 ms |  10.170 ms |  2.6412 ms |   351.8 ms |   358.4 ms |   352.9 ms |  15000.0000 |  6000.0000 | 2000.0000 |  64.71 MB |
| ImportFromXlsxBytesTest |     10000 |   758.2 ms |  10.212 ms |  2.6520 ms |   755.6 ms |   762.3 ms |   757.8 ms |  25000.0000 | 11000.0000 | 3000.0000 | 110.87 MB |
|  **ImportFromCsvBytesTest** |     **30000** |   **511.5 ms** |  **17.159 ms** |  **4.4562 ms** |   **504.0 ms** |   **514.9 ms** |   **512.7 ms** |  **31000.0000** |  **7000.0000** | **1000.0000** |  **109.1 MB** |
|  ImportFromXlsBytesTest |     30000 | 1,050.8 ms |  33.740 ms |  8.7622 ms | 1,040.2 ms | 1,064.3 ms | 1,049.8 ms |  47000.0000 | 18000.0000 | 3000.0000 | 186.28 MB |
| ImportFromXlsxBytesTest |     30000 | 2,358.6 ms | 284.891 ms | 73.9854 ms | 2,228.2 ms | 2,405.3 ms | 2,378.8 ms |  67000.0000 | 23000.0000 | 4000.0000 |  331.5 MB |
|  **ImportFromCsvBytesTest** |     **50000** |   **860.7 ms** |  **11.026 ms** |  **2.8634 ms** |   **857.3 ms** |   **865.1 ms** |   **860.0 ms** |  **50000.0000** | **12000.0000** | **1000.0000** |    **182 MB** |
|  ImportFromXlsBytesTest |     50000 | 1,703.0 ms |  23.050 ms |  5.9861 ms | 1,695.9 ms | 1,709.7 ms | 1,703.3 ms |  76000.0000 | 28000.0000 | 2000.0000 | 315.85 MB |
| ImportFromXlsxBytesTest |     50000 | 3,869.8 ms | 141.111 ms | 36.6460 ms | 3,817.0 ms | 3,909.6 ms | 3,872.4 ms | 110000.0000 | 36000.0000 | 3000.0000 | 552.76 MB |
|  **ImportFromCsvBytesTest** |     **65535** | **1,144.7 ms** |  **65.693 ms** | **17.0602 ms** | **1,121.2 ms** | **1,162.6 ms** | **1,149.4 ms** |  **65000.0000** | **16000.0000** | **1000.0000** | **238.24 MB** |
|  ImportFromXlsBytesTest |     65535 | 2,324.6 ms |  23.299 ms |  6.0507 ms | 2,318.4 ms | 2,333.9 ms | 2,323.8 ms |  99000.0000 | 35000.0000 | 3000.0000 | 413.07 MB |
| ImportFromXlsxBytesTest |     65535 | 5,011.3 ms | 108.362 ms | 28.1414 ms | 4,977.1 ms | 5,054.4 ms | 5,012.6 ms | 145000.0000 | 48000.0000 | 3000.0000 | 723.66 MB |
