using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    /// <summary>
    /// Excel样式表
    /// </summary>
    public class StyleConfig
    {
        public List<NumberFormatItem> NumberFormats { get; set; } = new List<NumberFormatItem>();
        public List<CellXfItem> CellXfs { get; set; } = new List<CellXfItem>();
    }

    public class NumberFormatItem
    {
        public int NumFmtId { get; set; }

        public string FormatCode { get; set; }
    }

    public class CellXfItem
    {
        public int ApplyNumFmtId { get; set; }

        public bool ApplyNumberFormat { get; set; }
    }
}
