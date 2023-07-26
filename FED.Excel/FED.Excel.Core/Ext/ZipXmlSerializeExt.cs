using FED.Excel.Core.ExcelXmlModel;

using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace FED.Excel.Core.Ext
{
    public static class ZipXmlSerializeExt
    {
        public static bool IsEndElement(this XmlReader reader, string nodeName)
        {
            return reader.Name == nodeName && reader.NodeType == XmlNodeType.EndElement;
        }

        public static ShareStringsTable DeserializeShareStrings(this ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                var shareStringsTable = new ShareStringsTable();
                using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true, IgnoreProcessingInstructions = true, XmlResolver = null }))
                {
                    var i = 0;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("t"))
                        {
                            reader.Read();
                            var content = reader.Value;
                            shareStringsTable.AddItem(i++, content);
                        }
                        reader.Read();
                    }
                    return shareStringsTable;
                }
            }
        }

        public static StyleConfig DeserializeStyleConfig(this ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                var styleConfig = new StyleConfig();
                using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true, IgnoreProcessingInstructions = true, XmlResolver = null }))
                {
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("numFmts"))
                        {
                            while (reader.Read())
                            {
                                if (reader.IsStartElement("numFmt"))
                                {
                                    var numFmtId = reader.GetAttribute("numFmtId");
                                    var formatCode = reader.GetAttribute("formatCode");
                                    styleConfig.NumberFormats.Add(new NumberFormatItem { NumFmtId = int.Parse(numFmtId), FormatCode = formatCode });
                                }
                                else if (reader.IsEndElement("numFmts"))
                                    break;
                            }
                        }
                        else if (reader.IsStartElement("cellXfs"))
                        {
                            while (reader.Read())
                            {
                                if (reader.IsStartElement("xf"))
                                {
                                    var numFmtId = reader.GetAttribute("numFmtId");
                                    var applyNumberFormat = reader.GetAttribute("applyNumberFormat");
                                    styleConfig.CellXfs.Add(new CellXfItem { ApplyNumFmtId = int.Parse(numFmtId), ApplyNumberFormat = applyNumberFormat == "1" });
                                }
                                else if (reader.IsEndElement("cellXfs"))
                                    break;
                            }
                        }
                        reader.Read();
                    }
                    return styleConfig;
                }
            }
        }


        public static SheetData DeserializeSheet(this ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                var sheetData = new SheetData();
                using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true, IgnoreProcessingInstructions = true, XmlResolver = null }))
                {
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("row"))
                        {
                            var rowNumber = reader.GetAttribute("r");
                            var row = new SheetRow { RowNumber = int.Parse(rowNumber) };
                            var cell = new SheetRowCell();
                            while (reader.Read())
                            {
                                if (reader.IsStartElement("c"))
                                {
                                    var cellNumber = reader.GetAttribute("r");
                                    var styleIdStr = reader.GetAttribute("s");
                                    int? styleId;
                                    if (int.TryParse(styleIdStr, out var value))
                                        styleId = value;
                                    else
                                        styleId = null;

                                    var cellType = reader.GetAttribute("t");
                                    cell = new SheetRowCell
                                    {
                                        CellNumber = cellNumber,
                                        CellType = cellType,
                                        StyleId = styleId
                                    };
                                    row.Cells.Add(cell);
                                }
                                else if (reader.IsStartElement("v"))
                                {
                                    reader.Read();
                                    var value = reader.Value;
                                    cell.Value = value;
                                }
                                else if (reader.IsEndElement("row"))
                                {
                                    sheetData.Rows.Add(row);
                                    break;
                                }
                            }
                        }
                        else if (reader.IsEndElement("sheetData"))
                            break;
                        reader.Read();
                    }
                    return sheetData;
                }
            }
        }
    }
}
