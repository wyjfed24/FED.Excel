using System.Collections.Generic;
using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    [XmlRoot(ElementName = "styleSheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class StyleXmlModel
    {
        [XmlElement("numFmts")]
        public NumberFormatInfo NumberFormatInfo { get; set; }

        [XmlElement("cellXfs")]
        public CellXfInfo CellXfInfo { get; set; }
    }

    public class NumberFormatInfo
    {
        [XmlElement("numFmt")]
        public List<NumberFormat> NumberFormats { get; set; }
    }

    public class NumberFormat
    {
        [XmlAttribute("numFmtId")]
        public int NumFmtId { get; set; }

        [XmlAttribute("formatCode")]
        public string FormatCode { get; set; }
    }
    public class CellXfInfo
    {
        [XmlElement("xf")]
        public List<CellXf> CellXfs { get; set; }

    }

    public class CellXf
    {
        [XmlAttribute("numFmtId")]
        public int NumFmtId { get; set; }

        [XmlAttribute("applyNumberFormat")]
        public bool ApplyNumberFormat { get; set; }
    }
}
