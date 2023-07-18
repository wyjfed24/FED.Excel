using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    [XmlRoot(ElementName = "worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class SheetXmlModel
    {
        /// <summary>
        /// Sheet名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Sheet数据
        /// </summary>
        [XmlElement("sheetData")]
        public SheetDataXmlModel SheetData { get; set; }
    }

    public class SheetDataXmlModel
    {
        /// <summary>
        /// 行集合
        /// </summary>
        [XmlElement("row")]
        public SheetDataRowXmlModel[] Rows { get; set; }
    }

    public class SheetDataRowXmlModel
    {
        /// <summary>
        /// 行号
        /// </summary>
        [XmlAttribute("r")]
        public int RowNumber { get; set; }

        /// <summary>
        /// 单元格集合
        /// </summary>
        [XmlElement("c")]
        public SheetDataRowColXmlModel[] Cells { get; set; }
    }

    public class SheetDataRowColXmlModel
    {
        /// <summary>
        /// 单元格标识
        /// </summary>
        [XmlAttribute("r")]
        public string CellNumber { get; set; }

        /// <summary>
        /// 样式表Id
        /// </summary>
        [XmlAttribute("s")]
        public int StyleId { get; set; }

        /// <summary>
        /// 单元格类型 “s”表示字符串，并且Value值为共享字符串表索引，空值表示Value存储实际值
        /// </summary>
        [XmlAttribute("t")]
        public string CellType { get; set; }

        /// <summary>
        /// 单元格值或共享字符串表索引
        /// </summary>
        [XmlElement("v")]
        public string Value { get; set; }
    }
}