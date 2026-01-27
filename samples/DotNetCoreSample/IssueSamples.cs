using WeihanLi.Extensions;

public static partial class IssueSamples
{
    public static void Issue169Sample()
    {
        var filePath = @"C:\Users\Weiha\Downloads\test\2.xlsx";
        var workbook = ExcelHelper.LoadExcel(
            File.OpenRead(filePath),
            ExcelFormat.Xlsx
        );
        var settings = FluentSettings.For<MaterielDetailDto>();
        settings.WithPostImportAction((x, rowIndex) => x?.RowNum = rowIndex + 1);
        var list = workbook.ToEntityList<MaterielDetailDto>();
        foreach (var item in list)
        {
            Console.WriteLine(item.ToJson());
        }
    }
}

[Sheet(SheetIndex = 0, StartRowIndex = 6, DisableColumnIndexAdjustment = true)]
public class MaterielDetailDto
{
    [Column(IsIgnored = true)]
    public int RowNum { get; set; }
    
    /// <summary>
    /// 编号
    /// </summary>
    [Column("物料编号", Index = 0)]
    public string No { get; set; }

    /// <summary>
    /// 物料名称
    /// </summary>
    [Column("物料名称", Index = 1)]
    public string Name { get; set; }

    /// <summary>
    /// 规格型号
    /// </summary>
    [Column("规格型号", Index = 2)]
    public string Specification { get; set; }

    /// <summary>+
    /// 特殊库存
    /// </summary>
    [Column("特殊库存", Index = 3)]
    public string QuantityM { get; set; }

    /// <summary>
    /// 单位
    /// </summary>
    [Column("计量单位", Index = 4)]
    public string Unit { get; set; }

    [Column("会计年度", Index = 5)]
    public int Ytime { get; set; }

    [Column("会计期间", Index = 6)]
    public int Mtime { get; set; }

    // 期初余额
    [Column("方向", Index = 7)]
    public string Direction1 { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量", Index = 8)]
    public decimal Quantity1 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价", Index = 9)]
    public decimal UnitPrice1 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额", Index = 10)]
    public decimal TotalAmount1 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价", Index = 11)]
    public decimal StandardUnitPrice1 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额", Index = 12)]
    public decimal StandardotalAmount1 { get; set; }

    // 期初余额

    //本期借方
    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量", Index = 13)]
    public decimal Quantity2 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价", Index = 14)]
    public decimal UnitPrice2 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额", Index = 15)]
    public decimal TotalAmount2 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价", Index = 16)]
    public decimal StandardUnitPrice2 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额", Index = 17)]
    public decimal StandardotalAmount2 { get; set; }

    //本期贷方

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量", Index = 18)]
    public decimal Quantity3 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价", Index = 19)]
    public decimal UnitPrice3 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额", Index = 20)]
    public decimal TotalAmount3 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价", Index = 21)]
    public decimal StandardUnitPrice3 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额", Index = 22)]
    public decimal StandardotalAmount3 { get; set; }

    //期末余额
    [Column("方向", Index = 23)]
    public string Direction2 { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量", Index = 24)]
    public decimal Quantity4 { get; set; }

    /// <summary>
    /// 实际单价
    /// </summary>
    [Column("实际单价", Index = 25)]
    public decimal UnitPrice4 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额", Index = 26)]
    public decimal TotalAmount4 { get; set; }

    [Column("核算类别名称", Index = 27)]
    public string CategoryName { get; set; }

    [Column("是否可用", Index = 28)]
    public string IsStop { get; set; }
}
