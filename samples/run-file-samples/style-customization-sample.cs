using NPOI.SS.UserModel;
using NPOI.SS.Util;

FluentSettings.LoadMappingProfile<StyledEntity, StyledEntityProfile>();

var list = new List<StyledEntity>
{
    new StyledEntity { Id = 1, Name = "Alice", Amount = 1500.50m, Status = "Approved", Date = DateTime.Now.AddDays(-10) },
    new StyledEntity { Id = 2, Name = "Bob", Amount = -200.75m, Status = "Pending", Date = DateTime.Now.AddDays(-5) },
    new StyledEntity { Id = 3, Name = "Charlie", Amount = 300.00m, Status = "Rejected", Date = DateTime.Now.AddDays(-2) },
    new StyledEntity { Id = 4, Name = "Diana", Amount = 450.25m, Status = "Approved", Date = DateTime.Now.AddDays(-1) },
};
const string path = @"C:\Users\Weiha\Downloads\test\styled-report.xlsx";
list.ToExcelFile(path);
Console.WriteLine($"Excel file generated at: {path}");

public class StyledEntity
{
    public int Id { get; set; }
    public required string Name { get; set; }
    public decimal Amount { get; set; }
    public required string Status { get; set; }
    public DateTime Date { get; set; }
}

public class StyledEntityProfile : IMappingProfile<StyledEntity>
{
    public void Configure(IExcelConfiguration<StyledEntity> configuration)
    {
        configuration.HasAuthor("Spark")
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
                    headerFont.FontName = "JetBrains Mono";
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
