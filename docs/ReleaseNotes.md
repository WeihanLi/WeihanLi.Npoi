# WeihanLi.Npoi Release Notes

## [1.13.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.13.0)

- add support for `EntityList`/`DataTable` export auto split sheets when needed

## [1.12.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.12.0)

- refact `ExcelSetting` and `SheetSetting`
- add support for `RowFilter` and `CellFilter`(mainly for import)
- add support for reading file when file opened by another process

## [1.11.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.11.0)

- add support for formula value import

## [1.10.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.10.0)

- add `EndRowIndex` for `SheetSetting`(zero-based, included)
- add FluentAPI `WithDataValidation` for excel setting, if set will ignore invalid data when import excel
- remove `CSVHelper` `TEntity` `new()` constraint

## [1.9.6](https://www.nuget.org/packages/WeihanLi.Npoi/1.9.6)

- fix xlsx workbook `AppVersion` property value caused warning

## [1.9.5](https://www.nuget.org/packages/WeihanLi.Npoi/1.9.5)

- fix `ExcelHelper.ToDataTable` bug with blank cell, thanks for hokis's feedback

## [1.9.4](https://www.nuget.org/packages/WeihanLi.Npoi/1.9.4)

- expose `CsvHelper.GetCsvText` extensions

## [1.9.2](https://www.nuget.org/packages/WeihanLi.Npoi/1.9.2)

- fix `CsvHelper.ParseLine` bug with quoted value, thanks for hokis's effort

## [1.9.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.9.0)

- remove `return 1` code fix #64
- optimize fluent formatter performance
- add `FluentSettings.For` instead of `ExcelHelper.SettingsFor`, fluent settings is not only for excel, but also work with csv

## [1.8.2](https://www.nuget.org/packages/WeihanLi.Npoi/1.8.2)

- add `TemplateHelper.ConfigureTemplateOptions` to allow user config templateParamFormat

## [1.8.1](https://www.nuget.org/packages/WeihanLi.Npoi/1.8.1)

- add `ExportExcelByTemplate`, fix #33
- update `NpoiRowCollection`
- optimize `DataTable` support for csv
- export csv as utf8 encoding
- update export generic type constraint, remove `new()` constraint

## [1.7.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.7.0)

- add `HasColumnInputFormatter`/`HasColumnOutputFormatter`
- simply `IExcelConfiguration` SheetConfiguration

## [1.6.1](https://www.nuget.org/packages/WeihanLi.Npoi/1.6.1)

- fix inherit property configure bug
- fix empty column skipped bug, fix with `row.Cells.Count` => `row.LastCellNumber`
- optimize `AdjustColumnIndex`
- allow use `Ignored(false)` to unignore property

## [1.6.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.6.0)

- add shadow property support
- add version info when export `*.xlsx` excel

## [1.5.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.5.0)

- add support for more format, treat as xlsx
- add `AutoColumnWidthEnabled` setting to `SheetSetting`, no autoSizeColumn by default
- add `CsvHelper.ToEntityList(byte[] bytes)`/`CsvHelper.ToEntityList(Stream stream)`
- use xls for default ExcelFormat(better performance)

## [1.4.5](https://www.nuget.org/packages/WeihanLi.Npoi/1.4.5)

- try to auto adjust column index when import excel(do not update existing settings)
- add `InputFormatter`/`OutputFormatter`
- apply column settings for CSV
- remove unused SheetConfiguration

## [1.4.4](https://www.nuget.org/packages/WeihanLi.Npoi/1.4.4)

- add `ExcelHelper.LoadExcel()`/`ExcelHelper.ToEntityList` override for stream/bytes

## [1.4.3](https://www.nuget.org/packages/WeihanLi.Npoi/1.4.3)

- fix possible `NullReferenceException` when `ExcelHelper.ToEntityList()`/`ToExcelFile()`
- fix treat `string.Empty` as `null` bug, `SetCellType` after `SetCellValue` so that `null` => `CellType.Blank`, `string.Empty` => `CellType.String`

## [1.4.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.4.0)

- add support for custom column width
- fix `ToExcelFile`/`ImportData` extension not applied configuration bug
- add support for specific sheetIndex when export excel

## [1.3.8](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.8)

- fix : `CsvHelper.ToDataTable()` and export DataTable to csv (Thanks for Arek's feedback)
- add support for no header when export excel(to fix #26 )

## [1.3.7](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.7)

- add `HasColumnFormatter<TEntity, TProperty>(Func<TEntity, TProperty, object> columnFormatter)` to support for custom column output for export

## [1.3.6](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.6)

- add [sourcelink](http://github.com/dotnet/sourcelink) support

## [1.3.5](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.5)

- add support for csv escape

## [1.3.3](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.3)

- fix csv custom columnIndex bug
- Optimize csv operation for entity list

## [1.3.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.3.0)

- Update NPOI package to 2.4.1
- Add support for struct types
- Add default excel settings

## [1.2.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.2.0)

- Update NPOI package to 2.4.0, use NPOI for netstandard2.0 also
- Add CsvHelper for import and export csv file, and mapping to entities

## [1.1.0](https://www.nuget.org/packages/WeihanLi.Npoi/1.1.0)

- StrongNaming package
