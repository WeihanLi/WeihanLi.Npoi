# 自定义 Excel 样式

## 简介

WeihanLi.Npoi 通过 sheet 配置中的 `RowAction` 和 `CellAction` 回调提供了强大的样式定制功能。你可以自定义字体、颜色、对齐方式、边框，并添加数据验证，以创建专业外观的 Excel 文件。

## 自动列宽

最简单的样式功能是启用自动列宽调整：

```csharp
var settings = FluentSettings.For<TestEntity>();
settings.HasSheetSetting(config =>
{
    config.AutoColumnWidthEnabled = true;
});
```

这会根据内容自动调整列宽，使你的电子表格更易读，无需手动调整宽度。

## 样式化标题行

使用 `RowAction` 自定义行的外观。一个常见的用例是样式化标题行：

```csharp
settings.HasSheetSetting(config =>
{
    config.StartRowIndex = 1;
    config.SheetName = "StyledSheet";
    config.AutoColumnWidthEnabled = true;

    config.RowAction = row =>
    {
        if (row.RowNum == 0) // 标题行
        {
            // 创建单元格样式
            var style = row.Sheet.Workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            
            // 创建和配置字体
            var font = row.Sheet.Workbook.CreateFont();
            font.FontName = "Arial";
            font.IsBold = true;
            font.FontHeight = 220; // 11pt (字体高度以 1/20 磅为单位)
            font.Color = IndexedColors.White.Index;
            
            style.SetFont(font);
            
            // 设置背景颜色
            style.FillForegroundColor = IndexedColors.DarkBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
            
            // 将样式应用于行中的所有单元格
            row.Cells.ForEach(c => c.CellStyle = style);
        }
    };
});
```

### 字体属性

可用的字体属性包括：

- `FontName`: 字体族（例如："宋体"、"微软雅黑"、"Arial"）
- `FontHeight`: 字体大小，以 1/20 磅为单位（例如：200 = 10pt, 220 = 11pt）
- `IsBold`: 粗体文本
- `IsItalic`: 斜体文本
- `IsStrikeout`: 删除线文本
- `Underline`: 文本下划线样式
- `Color`: 使用 `IndexedColors` 设置字体颜色

### 单元格样式属性

主要的单元格样式属性：

- `Alignment`: 水平对齐（`Left`、`Center`、`Right`、`Justify`）
- `VerticalAlignment`: 垂直对齐（`Top`、`Center`、`Bottom`）
- `FillForegroundColor`: 背景颜色
- `FillPattern`: 填充模式（通常使用 `SolidForeground` 实现纯色）
- `BorderTop`、`BorderBottom`、`BorderLeft`、`BorderRight`: 边框样式
- `WrapText`: 启用文本换行

## 样式化单个单元格

使用 `CellAction` 根据单元格位置或内容自定义单个单元格：

```csharp
settings.HasSheetSetting(config =>
{
    config.CellAction = cell =>
    {
        // 样式化特定列
        if (cell.ColumnIndex == 0) // 第一列
        {
            var style = cell.Sheet.Workbook.CreateCellStyle();
            var font = cell.Sheet.Workbook.CreateFont();
            font.IsBold = true;
            style.SetFont(font);
            cell.CellStyle = style;
        }
        
        // 基于内容的条件样式
        if (cell.RowIndex > 0 && cell.ColumnIndex == 3) // 数据行，第 4 列
        {
            if (cell.NumericCellValue < 0) // 负数显示为红色
            {
                var style = cell.Sheet.Workbook.CreateCellStyle();
                var font = cell.Sheet.Workbook.CreateFont();
                font.Color = IndexedColors.Red.Index;
                style.SetFont(font);
                cell.CellStyle = style;
            }
        }
    };
});
```

## 数据验证

为单元格添加下拉列表和其他验证规则：

```csharp
settings.HasSheetSetting(config =>
{
    config.CellAction = cell =>
    {
        // 为匹配"状态"的标题单元格添加下拉验证
        if (cell.RowIndex == 0 && cell.StringCellValue == "状态")
        {
            var validationHelper = cell.Sheet.GetDataValidationHelper();
            
            // 定义允许的值
            var statusValues = new[] { "活跃", "停用", "待定" };
            var constraint = validationHelper.CreateExplicitListConstraint(statusValues);
            
            // 将验证应用于数据行（第 1-100 行，当前列）
            var addressList = new CellRangeAddressList(1, 100, cell.ColumnIndex, cell.ColumnIndex);
            var validation = validationHelper.CreateValidation(constraint, addressList);
            
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("状态无效", "请从下拉列表中选择");
            validation.ShowPromptBox = true;
            validation.CreatePromptBox("状态选择", "从列表中选择一个状态");
            
            cell.Sheet.AddValidationData(validation);
        }
    };
});
```

### 验证类型

可用的不同验证类型：

```csharp
// 列表验证（下拉列表）
var constraint = validationHelper.CreateExplicitListConstraint(new[] { "选项1", "选项2" });

// 整数验证
var intConstraint = validationHelper.CreateIntegerConstraint(
    OperatorType.Between, "1", "100");

// 小数验证
var decimalConstraint = validationHelper.CreateDecimalConstraint(
    OperatorType.GreaterThan, "0", null);

// 日期验证
var dateConstraint = validationHelper.CreateDateConstraint(
    OperatorType.Between, "2024-01-01", "2024-12-31", "yyyy-MM-dd");

// 文本长度验证
var textConstraint = validationHelper.CreateTextLengthConstraint(
    OperatorType.LessThan, "100", null);
```

## 完整示例

这是一个结合多种样式功能的综合示例：

```csharp
public class StyledEntityProfile : IMappingProfile<StyledEntity>
{
    public void Configure(IExcelConfiguration<StyledEntity> configuration)
    {
        configuration.HasAuthor("您的名字")
            .HasTitle("样式化报表")
            .HasDescription("专业样式的 Excel 报表");

        configuration.HasSheetSetting(config =>
        {
            config.SheetName = "报表";
            config.StartRowIndex = 1;
            config.AutoColumnWidthEnabled = true;

            // 样式化标题行
            config.RowAction = row =>
            {
                if (row.RowNum == 0)
                {
                    var headerStyle = row.Sheet.Workbook.CreateCellStyle();
                    headerStyle.Alignment = HorizontalAlignment.Center;
                    headerStyle.VerticalAlignment = VerticalAlignment.Center;
                    headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
                    headerStyle.FillPattern = FillPattern.SolidForeground;
                    
                    var headerFont = row.Sheet.Workbook.CreateFont();
                    headerFont.FontName = "微软雅黑";
                    headerFont.IsBold = true;
                    headerFont.FontHeight = 240; // 12pt
                    headerStyle.SetFont(headerFont);
                    
                    // 添加边框
                    headerStyle.BorderBottom = BorderStyle.Thin;
                    headerStyle.BorderTop = BorderStyle.Thin;
                    headerStyle.BorderLeft = BorderStyle.Thin;
                    headerStyle.BorderRight = BorderStyle.Thin;
                    
                    row.Cells.ForEach(c => c.CellStyle = headerStyle);
                }
            };

            // 添加验证和条件格式
            config.CellAction = cell =>
            {
                // 为状态列添加验证
                if (cell.RowIndex == 0 && cell.StringCellValue == "状态")
                {
                    var validationHelper = cell.Sheet.GetDataValidationHelper();
                    var statusList = new[] { "已批准", "待定", "已拒绝" };
                    var constraint = validationHelper.CreateExplicitListConstraint(statusList);
                    var addressList = new CellRangeAddressList(1, 1000, cell.ColumnIndex, cell.ColumnIndex);
                    var validation = validationHelper.CreateValidation(constraint, addressList);
                    validation.ShowErrorBox = true;
                    cell.Sheet.AddValidationData(validation);
                }
                
                // 用红色突出显示负数金额
                if (cell.RowIndex > 0 && cell.ColumnIndex == 2) // 金额列
                {
                    try
                    {
                        if (cell.NumericCellValue < 0)
                        {
                            var redStyle = cell.Sheet.Workbook.CreateCellStyle();
                            var redFont = cell.Sheet.Workbook.CreateFont();
                            redFont.Color = IndexedColors.Red.Index;
                            redFont.IsBold = true;
                            redStyle.SetFont(redFont);
                            cell.CellStyle = redStyle;
                        }
                    }
                    catch { } // 如果不是数值单元格则跳过
                }
            };
        });

        // 配置属性
        configuration.Property(x => x.Id).HasColumnIndex(0);
        configuration.Property(x => x.Name).HasColumnIndex(1);
        configuration.Property(x => x.Amount).HasColumnIndex(2);
        configuration.Property(x => x.Status).HasColumnIndex(3);
        configuration.Property(x => x.Date)
            .HasColumnIndex(4)
            .HasColumnFormatter("yyyy-MM-dd");
    }
}
```

## 最佳实践

1. **重用样式**：创建样式一次并重用，而不是为每个单元格创建新样式，以提高性能并减小文件大小。

2. **性能考虑**：在处理大型数据集时应用样式时要注意性能。考虑仅样式化标题行或特定列。

3. **颜色一致性**：使用 `IndexedColors` 以在不同的 Excel 版本之间保持一致的颜色。

4. **字体大小**：记住字体高度以 1/20 磅为单位（将磅值乘以 20）。

5. **验证范围**：为数据验证设置适当的范围以覆盖预期的数据行。

## 参考

- [示例实现](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [NPOI 文档](http://poi.apache.org/)
