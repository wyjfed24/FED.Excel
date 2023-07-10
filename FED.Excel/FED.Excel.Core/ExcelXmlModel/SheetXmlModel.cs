using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    [XmlRoot(ElementName = "worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class SheetXmlModel
    {
        [XmlElement("sheetData")]
        public SheetDataXmlModel SheetData { get; set; }
    }

    public class SheetDataXmlModel
    {

        [XmlElement("row")]
        public SheetDataRowXmlModel[] Rows { get; set; }
    }

    public class SheetDataRowXmlModel
    {
        [XmlAttribute("r")]
        public int RowNumber { get; set; }

        [XmlElement("c")]
        public SheetDataRowColXmlModel[] Cells { get; set; }
    }

    public class SheetDataRowColXmlModel
    {
        [XmlAttribute("r")]
        public string CellNumber { get; set; }

        [XmlAttribute("s")]
        public int StyleId { get; set; }

        [XmlAttribute("t")]
        public string CellType { get; set; }

        [XmlElement("v")]
        public string Value { get; set; }
    }
}