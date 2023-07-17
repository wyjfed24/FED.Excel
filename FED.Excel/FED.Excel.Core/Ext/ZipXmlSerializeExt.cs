using System.IO.Compression;
using System.Xml;
using System.Xml.Serialization;

namespace FED.Excel.Core.Ext
{
    public static class ZipXmlSerializeExt
    {
        public static XmlModel Deserialize<XmlModel>(this ZipArchiveEntry entry) where XmlModel : class
        {
            using (var stream = entry.Open())
            {
                using (var reader = XmlReader.Create(stream))
                {
                    var xmlSerializer = new XmlSerializer(typeof(XmlModel));
                    var obj = xmlSerializer.Deserialize(reader) as XmlModel;
                    return obj;
                }
            }
        }
    }
}
