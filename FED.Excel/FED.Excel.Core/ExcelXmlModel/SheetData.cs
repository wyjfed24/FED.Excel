using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    public class SheetData
    {
        public string Name { get; set; }

        public List<SheetRow> Rows { get; set; } = new List<SheetRow>();
    }

    public class SheetRow
    {
        public int RowNumber { get; set; }

        public List<SheetRowCell> Cells { get; set; } = new List<SheetRowCell>();
    }

    public class SheetRowCell
    {
        /// <summary>
        /// 单元格标识
        /// </summary>
        public string CellNumber { get; set; }

        /// <summary>
        /// 样式表Id
        /// </summary>
        public int StyleId { get; set; }

        /// <summary>
        /// 单元格类型 “s”表示字符串，并且Value值为共享字符串表索引，空值表示Value存储实际值
        /// </summary>
        public string CellType { get; set; }

        /// <summary>
        /// 单元格值或共享字符串表索引
        /// </summary>
        public string Value { get; set; }
    }
}
