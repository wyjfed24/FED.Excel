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
                        if (reader.Name == "t")
                        {
                            reader.MoveToContent();
                            var content = reader.ReadElementContentAsString();
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
                        if (reader.Name == "numFmts" && reader.NodeType == XmlNodeType.Element)
                        {
                            while (reader.Read())
                            {
                                if (reader.Name == "numFmt" && reader.NodeType == XmlNodeType.Element)
                                {
                                    var numFmtId = reader.GetAttribute("numFmtId");
                                    var formatCode = reader.GetAttribute("formatCode");
                                    styleConfig.NumberFormats.Add(new NumberFormatItem { NumFmtId = int.Parse(numFmtId), FormatCode = formatCode });
                                }
                                else if (reader.Name == "numFmts" && reader.NodeType == XmlNodeType.EndElement)
                                    break;
                            }
                        }
                        else if (reader.Name == "cellXfs" && reader.NodeType == XmlNodeType.Element)
                        {
                            while (reader.Read())
                            {
                                if (reader.Name == "xf" && reader.NodeType == XmlNodeType.Element)
                                {
                                    var numFmtId = reader.GetAttribute("numFmtId");
                                    var applyNumberFormat = reader.GetAttribute("applyNumberFormat");
                                    styleConfig.CellXfs.Add(new CellXfItem { ApplyNumFmtId = int.Parse(numFmtId), ApplyNumberFormat = applyNumberFormat == "1" });
                                }
                                else if (reader.Name == "cellXfs" && reader.NodeType == XmlNodeType.EndElement)
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
                        if (reader.Name == "row" && reader.NodeType == XmlNodeType.Element)
                        {
                            var rowNumber = reader.GetAttribute("r");
                            var row = new SheetRow { RowNumber = int.Parse(rowNumber) };
                            var cell=new SheetRowCell();
                            while (reader.Read())
                            {
                                if (reader.Name == "c" && reader.NodeType == XmlNodeType.Element)
                                {
                                    var cellNumber = reader.GetAttribute("r");
                                    var styleId = reader.GetAttribute("s");
                                    var cellType = reader.GetAttribute("t");
                                    cell = new SheetRowCell
                                    {
                                        CellNumber = cellNumber,
                                        CellType = cellType,
                                        StyleId = int.Parse(styleId)
                                    };
                                    row.Cells.Add(cell);
                                }
                                else if (reader.Name == "v" && reader.NodeType == XmlNodeType.Element)
                                {
                                    reader.Read();
                                    var value = reader.Value;
                                    cell.Value = value;
                                }
                                else if (reader.Name == "row" && reader.NodeType == XmlNodeType.EndElement)
                                {
                                    sheetData.Rows.Add(row);
                                    break;
                                }
                            }
                        }
                        else if (reader.Name == "sheetData" && reader.NodeType == XmlNodeType.EndElement)
                            break;
                        reader.Read();
                    }
                    return sheetData;
                }
            }
        }
    }
}
