using FED.Excel.Core.Attributes;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FED.Excel.Test
{
    public class BigTest
    {
        [ExcelColumn("A")]
        public string A { get; set; }

        [ExcelColumn("B")]
        public string B { get; set; }

        [ExcelColumn("C")]
        public string C { get; set; }

        [ExcelColumn("D")]
        public string D { get; set; }
        
        [ExcelColumn("E")]
        public string E { get; set; }

        [ExcelColumn("F")]
        public string F { get; set; }

        [ExcelColumn("G")]
        public string G { get; set; }

        [ExcelColumn("H")]
        public string H { get; set; }

        [ExcelColumn("I")]
        public string I { get; set; }

        [ExcelColumn("J")]
        public string J { get; set; }
    }
    public class Sheet01
    {
        [ExcelColumn("当前章节")]
        public string No { get; set; }

        [ExcelColumn("所属章节")]
        public string ParentNo { get; set; }

        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("是否不计价目录")]
        public string IsCategory { get; set; }
    }

    public class Sheet02
    {
        [ExcelColumn("所属章节")]
        public string ParentNo { get; set; }

        [ExcelColumn("类型（设备、主材）")]
        public EnumTest01 Type { get; set; }

        [ExcelColumn("编号")]
        public string No { get; set; }

        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("规格")]
        public string Pack { get; set; }

        [ExcelColumn("单位")]
        public string Unit { get; set; }

        [ExcelColumn("含税价（元）")]
        public decimal? TaxPrice { get; set; }

        [ExcelColumn("不含税价（元）")]
        public decimal? NoTaxPrice { get; set; }

        [ExcelColumn("供货方")]
        public EnumTest04 Support { get; set; }

        [ExcelColumn("运杂费率（%）")]
        public decimal? TranRate { get; set; }

        [ExcelColumn("损耗率（%）")]
        public decimal? BadRate { get; set; }

        [ExcelColumn("包装系数（%）")]
        public decimal? PackageRate { get; set; }

        [ExcelColumn("单重(kg)")]
        public decimal? Weight { get; set; }

        [ExcelColumn("类别/运输类型")]
        public string TranType { get; set; }

        [ExcelColumn("材料类型(混凝土、配合比材料)")]
        public EnumTest03? MaterialType { get; set; }
    }
    public enum EnumTest01
    {
        主材 = 1,
        设备 = 2,
        土建 = 3
    }

    public enum EnumTest02
    {
        商品混凝土 = 1
    }

    public enum EnumTest03
    {
        混凝土 = 1,
        配合比材料 = 2
    }
    public enum EnumTest04
    {
        甲供 = 1,
        乙供 = 2
    }
}
