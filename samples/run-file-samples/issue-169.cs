var filePath = @"C:\Users\Weiha\Downloads\test\2.xlsx";
var workbook = ExcelHelper.LoadExcel(
     File.OpenRead(filePath),
     ExcelFormat.Xlsx
);
var settings = FluentSettings.For<MaterielDetailDto>();
settings.Property(x => x.RowNum).HasCellReader(c => c.RowIndex);
var list = workbook.ToEntityList<MaterielDetailDto>();
foreach (var item in list)
{
    Console.WriteLine($"#{item.RowNum+1} => {item.No}\t{item.Name}\t{item.Specification}\t{item.QuantityM}\t{item.Unit}");
}

[Sheet(SheetIndex = 0, StartRowIndex = 6)]
public partial class MaterielDetailDto
{
    public int RowNum { get; set; }
    
    /// <summary>
    /// 编号
    /// </summary>
    [Column("物料编号")]
    public string No { get; set; }

    /// <summary>
    /// 物料名称
    /// </summary>
    [Column("物料名称")]
    public string Name { get; set; }

    /// <summary>
    /// 规格型号
    /// </summary>
    [Column("规格型号")]
    public string Specification { get; set; }

    /// <summary>+
    /// 特殊库存
    /// </summary>
    [Column("特殊库存")]
    public string QuantityM { get; set; }

    /// <summary>
    /// 单位
    /// </summary>
    [Column("计量单位")]
    public string Unit { get; set; }

    [Column("会计年度")]
    public int Ytime { get; set; }

    [Column("会计期间")]
    public int Mtime { get; set; }

    // 期初余额
    [Column("方向")]
    public string Direction1 { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量")]
    public decimal Quantity1 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价")]
    public decimal UnitPrice1 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额")]
    public decimal TotalAmount1 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价")]
    public decimal StandardUnitPrice1 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额")]
    public decimal StandardotalAmount1 { get; set; }

    // 期初余额

    //本期借方
    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量")]
    public decimal Quantity2 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价")]
    public decimal UnitPrice2 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额")]
    public decimal TotalAmount2 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价")]
    public decimal StandardUnitPrice2 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额")]
    public decimal StandardotalAmount2 { get; set; }

    //本期贷方

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量")]
    public decimal Quantity3 { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    [Column("实际单价")]
    public decimal UnitPrice3 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额")]
    public decimal TotalAmount3 { get; set; }

    /// <summary>
    /// 标准单价
    /// </summary>
    [Column("标准单价")]
    public decimal StandardUnitPrice3 { get; set; }

    /// <summary>
    /// 标准金额
    /// </summary>
    [Column("标准金额")]
    public decimal StandardotalAmount3 { get; set; }

    //期末余额
    [Column("方向")]
    public string Direction2 { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    [Column("数量")]
    public decimal Quantity4 { get; set; }

    /// <summary>
    /// 实际单价
    /// </summary>
    [Column("实际单价")]
    public decimal UnitPrice4 { get; set; }

    /// <summary>
    /// 总金额
    /// </summary>
    [Column("实际金额")]
    public decimal TotalAmount4 { get; set; }

    [Column("核算类别名称")]
    public string CategoryName { get; set; }

    [Column("是否可用")]
    public string IsStop { get; set; }
}
