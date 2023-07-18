using System.Xml.Serialization;

namespace FED.Excel.Core.ExcelXmlModel
{
    [XmlRoot(ElementName = "sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class SharedStringsTableXmlModel
    {
        /// <summary>
        /// 共享字符串集合
        /// </summary>
        [XmlElement("si")]
        public SharedStringXmlModel[] SharedString { get; set; }

        public string this[int index]
        {
            get { return SharedString[index].Text; }
        }
    }

    public class SharedStringXmlModel
    {
        /// <summary>
        /// 字符串值
        /// </summary>
        [XmlElement("t")]
        public string Text { get; set; }
    }
}