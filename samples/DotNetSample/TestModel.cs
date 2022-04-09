// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System;
using WeihanLi.Npoi.Attributes;

namespace DotNetSample;

internal class Model
{
    [Column("酒店编号", Index = 0)]
    public string? HotelId { get; set; }

    [Column("订单号", Index = 1)]
    public string? OrderNo { get; set; }

    [Column("酒店名称", Index = 2)]
    public string? HotelName { get; set; }

    [Column("客户名称", Index = 3)]
    public string? CustomerName { get; set; }

    [Column(nameof(房型名称), Index = 4)]
    public string? 房型名称 { get; set; }

    [Column(nameof(入住日期), Index = 5, Formatter = "yyyy/M/d")]
    public DateTime 入住日期 { get; set; }

    [Column(nameof(离店日期), Index = 6, Formatter = "yyyy/M/d")]
    public DateTime 离店日期 { get; set; }

    [Column(nameof(间夜), Index = 7)]
    public int 间夜 { get; set; }

    [Column(nameof(支付类型), Index = 8)]
    public string? 支付类型 { get; set; }

    [Column(nameof(订单金额), Index = 9)]
    public decimal 订单金额 { get; set; }

    [Column(nameof(佣金率), Index = 10)]
    public decimal 佣金率 { get; set; }

    [Column(nameof(服务费), Index = 11)]
    public decimal 服务费 { get; set; }
}
