# WeihanLi.Npoi [![WeihanLi.Npoi](https://img.shields.io/nuget/v/WeihanLi.Npoi.svg)](https://www.nuget.org/packages/WeihanLi.Npoi/)

## Build

[![Build Status](https://travis-ci.org/WeihanLi/WeihanLi.Npoi.svg?branch=master)](https://travis-ci.org/WeihanLi/WeihanLi.Npoi)

## Intro

Npoi 扩展,适用于.netframework4.5及以上和netstandard2.0, .netframework基于NPOI, .netstandard基于 DotNetCore.NPOI

## Use

### Install

.NetFramework

``` bash
Install-Package WeihanLi.Npoi
```

.NetCore

``` bash
dotnet add package WeihanLi.Npoi
```

### GetStarted

1. Attributes

    Add `ColumnAttribute` on the property of the entity which you used for export or import

    Add `SheetAttribute` on the entity which you used for export or import,you can set the `StartRowIndex` on your need(by default it is `1`)

    for example:

    ``` csharp
    public class TestEntity
    {
        [Column("PKID")]
        public int PKID { get; set; }

        [Column("账单标题")]
        public string BillTitle { get; set; }

        [Column("账单详情")]
        public string BillDetails { get; set; }

        [Column("创建人")]
        public string CreatedBy { get; set; }

        [Column("创建时间")]
        public DateTime CreatedTime { get; set; }
    }


    internal class TestEntity1
    {
        /// <summary>
        /// 用户名
        /// </summary>
        [Column("用户名")]
        public string Username { get; set; }

        [Column(IsIgnored = true)]
        public string PasswordHash { get; set; }

        [Column("可用余额")]
        public decimal Amount { get; set; } = 1000M;

        [Column("微信id")]
        public string WechatOpenId { get; set; }

        [Column("是否启用")]
        public bool IsActive { get; set; }
    }
    ```

1. FluentApi

    You can also use FluentApi above version `1.0.3`

    for example:

    ``` csharp
    var setting = ExcelHelper.SettingFor<TestEntity>();
    // ExcelSetting
    setting.HasAuthor("WeihanLi")
        .HasTitle("WeihanLi.Npoi")
        .HasDescription("")
        .HasSubject("");

    setting.Property(_ => _.Amount)
        .HasColumnTitle("可用余额")
        .HasColumnIndex(2);

    setting.Property(_ => _.Username)
        .HasColumnTitle("用户名")
        .HasColumnIndex(0);

    setting.Property(_ => _.PasswordHash)
        .Ignored();
    ```

### Contact

Contact me: weihanli@outlook.com