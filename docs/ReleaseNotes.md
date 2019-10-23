# WeihanLi.Npoi Release Notes

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
