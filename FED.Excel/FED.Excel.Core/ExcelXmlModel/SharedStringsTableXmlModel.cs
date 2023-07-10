using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    [XmlRoot(ElementName = "sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class SharedStringsTableXmlModel
    {
        [XmlElement("si")]
        public SharedStringXmlModel[] SharedString { get; set; }

        public string this[int index]
        {
            get { return SharedString[index].Text; }
        }
    }

    public class SharedStringXmlModel
    {
        [XmlElement("t")]
        public string Text { get; set; }
    }
}