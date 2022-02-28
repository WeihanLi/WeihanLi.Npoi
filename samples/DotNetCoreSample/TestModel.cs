// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the MIT license.

using System;
using WeihanLi.Npoi.Attributes;

namespace DotNetCoreSample;

[Sheet(SheetIndex = 0, SheetName = "Abc", StartRowIndex = 1, EndRowIndex = 10, StartColumnIndex = 1, EndColumnIndex = 19)]
internal class ppDto
{
    [Column("创建日期")]
    public DateTime? CreateDate { get; set; }

    [Column("体积")]
    public decimal? Volume { get; set; }

    [Column("总重量")]
    public decimal? TotalWeight { get; set; }

    [Column("出错")]
    public string? Error { get; set; }

    [Column("真假")]
    public Boolean? TrueOrFalse { get; set; }

    [Column("范围")]
    public string? Range { get; set; }

    [Column("从")]
    public long? CtnFm { get; set; }

    [Column("到")]
    public long? CtnTo { get; set; }

    [Column("开始序列号")]
    public long? SerialStart { get; set; }

    [Column("截止序列号")]
    public long? SerialEnd { get; set; }

    [Column("包装代码")]
    public string? PackCode { get; set; }

    [Column("行号")]
    public string? Row { get; set; }

    [Column("买方项目号")]
    public string? BuyerNo { get; set; }

    [Column("SKU号")]
    public string? SKUNo { get; set; }

    [Column("订单号码")]
    public string? TradingPO { get; set; }

    [Column("MAIN LINE #")]
    public string? Item { get; set; }

    [Column("Color Name")]
    public string? ColorCode { get; set; }

    [Column("Size")]
    public string? Size { get; set; }

    [Column("简短描述")]
    public string? ContractColor { get; set; }

    [Column("发货方式")]
    public string? ShipWay { get; set; }

    [Column("数量")]
    public int? TtlQty { get; set; }

    [Column("内部包装的项目数量")]
    public int? Qty { get; set; }

    [Column("内包装计数")]
    public int? RatioQty { get; set; }

    [Column("箱数")]
    public int? CntQty { get; set; }

    [Column("R")]
    public string? LastCarton { get; set; }

    [Column("外箱代码")]
    public string? CartonCode { get; set; }

    [Column("净净重")]
    public decimal? NNWeight { get; set; }

    [Column("净重")]
    public decimal? NWeight { get; set; }

    [Column("毛重")]
    public decimal? GWeight { get; set; }

    [Column("单位")]
    public string? Unit { get; set; }

    [Column("长")]
    public decimal? cartonL { get; set; }

    [Column("宽")]
    public decimal? cartonW { get; set; }

    [Column("高")]
    public decimal? cartonH { get; set; }

    [Column("单位2")]
    public string? Unit2 { get; set; }

    [Column("扫描ID")]
    public string? ScanID { get; set; }
}
