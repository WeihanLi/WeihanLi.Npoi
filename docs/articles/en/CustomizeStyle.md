# Customize Excel Styles

## Introduction

WeihanLi.Npoi provides powerful styling capabilities through the `RowAction` and `CellAction` callbacks in sheet configuration. You can customize fonts, colors, alignment, borders, and add data validation to create professional-looking Excel files.

## Auto Column Width

The simplest styling feature is enabling automatic column width adjustment:

```csharp
var settings = FluentSettings.For<TestEntity>();
settings.HasSheetSetting(config =>
{
    config.AutoColumnWidthEnabled = true;
});
```

This automatically adjusts column widths based on content, making your spreadsheet more readable without manual width adjustments.

## Styling Header Rows

Use `RowAction` to customize the appearance of rows. A common use case is styling the header row:

```csharp
settings.HasSheetSetting(config =>
{
    config.StartRowIndex = 1;
    config.SheetName = "StyledSheet";
    config.AutoColumnWidthEnabled = true;

    config.RowAction = row =>
    {
        if (row.RowNum == 0) // Header row
        {
            // Create a cell style
            var style = row.Sheet.Workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            
            // Create and configure font
            var font = row.Sheet.Workbook.CreateFont();
            font.FontName = "Arial";
            font.IsBold = true;
            font.FontHeight = 220; // 11pt (font height is in 1/20th of a point)
            font.Color = IndexedColors.White.Index;
            
            style.SetFont(font);
            
            // Set background color
            style.FillForegroundColor = IndexedColors.DarkBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
            
            // Apply style to all cells in the row
            row.Cells.ForEach(c => c.CellStyle = style);
        }
    };
});
```

### Font Properties

Available font properties include:

- `FontName`: Font family (e.g., "Arial", "Calibri", "Times New Roman")
- `FontHeight`: Font size in 1/20th of a point (e.g., 200 = 10pt, 220 = 11pt)
- `IsBold`: Bold text
- `IsItalic`: Italic text
- `IsStrikeout`: Strikethrough text
- `Underline`: Text underline style
- `Color`: Font color using `IndexedColors`

### Cell Style Properties

Key cell style properties:

- `Alignment`: Horizontal alignment (`Left`, `Center`, `Right`, `Justify`)
- `VerticalAlignment`: Vertical alignment (`Top`, `Center`, `Bottom`)
- `FillForegroundColor`: Background color
- `FillPattern`: Fill pattern (usually `SolidForeground` for solid colors)
- `BorderTop`, `BorderBottom`, `BorderLeft`, `BorderRight`: Border styles
- `WrapText`: Enable text wrapping

## Styling Individual Cells

Use `CellAction` to customize individual cells based on their position or content:

```csharp
settings.HasSheetSetting(config =>
{
    config.CellAction = cell =>
    {
        // Style specific columns
        if (cell.ColumnIndex == 0) // First column
        {
            var style = cell.Sheet.Workbook.CreateCellStyle();
            var font = cell.Sheet.Workbook.CreateFont();
            font.IsBold = true;
            style.SetFont(font);
            cell.CellStyle = style;
        }
        
        // Conditional styling based on content
        if (cell.RowIndex > 0 && cell.ColumnIndex == 3) // Data rows, 4th column
        {
            if (cell.NumericCellValue < 0) // Negative numbers in red
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

## Data Validation

Add dropdown lists and other validation rules to cells:

```csharp
settings.HasSheetSetting(config =>
{
    config.CellAction = cell =>
    {
        // Add dropdown validation for header cell matching "Status"
        if (cell.RowIndex == 0 && cell.StringCellValue == "Status")
        {
            var validationHelper = cell.Sheet.GetDataValidationHelper();
            
            // Define allowed values
            var statusValues = new[] { "Active", "Inactive", "Pending" };
            var constraint = validationHelper.CreateExplicitListConstraint(statusValues);
            
            // Apply validation to data rows (rows 1-100, current column)
            var addressList = new CellRangeAddressList(1, 100, cell.ColumnIndex, cell.ColumnIndex);
            var validation = validationHelper.CreateValidation(constraint, addressList);
            
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("Invalid Status", "Please select from the dropdown list");
            validation.ShowPromptBox = true;
            validation.CreatePromptBox("Status Selection", "Choose a status from the list");
            
            cell.Sheet.AddValidationData(validation);
        }
    };
});
```

### Validation Types

Different validation types are available:

```csharp
// List validation (dropdown)
var constraint = validationHelper.CreateExplicitListConstraint(new[] { "Option1", "Option2" });

// Integer validation
var intConstraint = validationHelper.CreateIntegerConstraint(
    OperatorType.Between, "1", "100");

// Decimal validation
var decimalConstraint = validationHelper.CreateDecimalConstraint(
    OperatorType.GreaterThan, "0", null);

// Date validation
var dateConstraint = validationHelper.CreateDateConstraint(
    OperatorType.Between, "2024-01-01", "2024-12-31", "yyyy-MM-dd");

// Text length validation
var textConstraint = validationHelper.CreateTextLengthConstraint(
    OperatorType.LessThan, "100", null);
```

## Complete Example

Here's a comprehensive example combining multiple styling features:

```csharp
public class StyledEntityProfile : IMappingProfile<StyledEntity>
{
    public void Configure(IExcelConfiguration<StyledEntity> configuration)
    {
        configuration.HasAuthor("Your Name")
            .HasTitle("Styled Report")
            .HasDescription("Professional styled Excel report");

        configuration.HasSheetSetting(config =>
        {
            config.SheetName = "Report";
            config.StartRowIndex = 1;
            config.AutoColumnWidthEnabled = true;

            // Style header row
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
                    headerFont.FontName = "Calibri";
                    headerFont.IsBold = true;
                    headerFont.FontHeight = 240; // 12pt
                    headerStyle.SetFont(headerFont);
                    
                    // Add borders
                    headerStyle.BorderBottom = BorderStyle.Thin;
                    headerStyle.BorderTop = BorderStyle.Thin;
                    headerStyle.BorderLeft = BorderStyle.Thin;
                    headerStyle.BorderRight = BorderStyle.Thin;
                    
                    row.Cells.ForEach(c => c.CellStyle = headerStyle);
                }
            };

            // Add validation and conditional formatting
            config.CellAction = cell =>
            {
                // Add validation for status column
                if (cell.RowIndex == 0 && cell.StringCellValue == "Status")
                {
                    var validationHelper = cell.Sheet.GetDataValidationHelper();
                    var statusList = new[] { "Approved", "Pending", "Rejected" };
                    var constraint = validationHelper.CreateExplicitListConstraint(statusList);
                    var addressList = new CellRangeAddressList(1, 1000, cell.ColumnIndex, cell.ColumnIndex);
                    var validation = validationHelper.CreateValidation(constraint, addressList);
                    validation.ShowErrorBox = true;
                    cell.Sheet.AddValidationData(validation);
                }
                
                // Highlight negative amounts in red
                if (cell.RowIndex > 0 && cell.ColumnIndex == 2) // Amount column
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
                    catch { } // Skip if not a numeric cell
                }
            };
        });

        // Configure properties
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

## Best Practices

1. **Reuse Styles**: Create styles once and reuse them rather than creating new styles for each cell to improve performance and reduce file size.

2. **Performance**: Be mindful of performance when applying styles to large datasets. Consider styling only header rows or specific columns.

3. **Color Consistency**: Use `IndexedColors` for consistent coloring across different Excel versions.

4. **Font Sizes**: Remember that font height is in 1/20th of a point (multiply point size by 20).

5. **Validation Ranges**: Set appropriate ranges for data validation to cover expected data rows.

## References

- [Sample Implementation](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [NPOI Documentation](http://poi.apache.org/)
