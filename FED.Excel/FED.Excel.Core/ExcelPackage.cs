using FED.Excel.Core.ExcelXmlModel;
using FED.Excel.Core.Ext;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;

namespace FED.Excel.Core
{
    internal class ExcelPackage : IDisposable
    {
        private ZipArchive _zip;
        /// <summary>
        /// 共享字符串表
        /// </summary>
        internal ShareStringsTable SharedStrings { get; set; }

        /// <summary>
        /// 样式表
        /// </summary>
        internal StyleConfig Style { get; set; }

        /// <summary>
        /// Sheet原始数据集合
        /// </summary>
        internal List<SheetData> Sheets { get; set; }

        internal ExcelPackage(Stream stream)
        {
            _zip = new ZipArchive(stream);
            Build();
        }

        public void Build()
        {
            SharedStrings = GetShareStrings(_zip);
            Style = GetStyles(_zip);
            Sheets = GetSheetDatas(_zip);
        }

        #region Excel文件解析

        /// <summary>
        /// 获取共享字符串表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private ShareStringsTable GetShareStrings(ZipArchive zip)
        {
            var sharedStringsEntry = zip.Entries.Where(x => x.FullName == "xl/sharedStrings.xml").FirstOrDefault();
            if (sharedStringsEntry == null)
                return new ShareStringsTable();
            var sharedStringsTable = sharedStringsEntry.DeserializeShareStrings();
            return sharedStringsTable;
        }

        /// <summary>
        /// 获取样式表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private StyleConfig GetStyles(ZipArchive zip)
        {
            var styleEntry = zip.Entries.Where(x => x.FullName == "xl/styles.xml").FirstOrDefault();
            if (styleEntry == null)
                return new StyleConfig();
            var style = styleEntry.DeserializeStyleConfig();
            return style;
        }

        /// <summary>
        /// 获取Sheet数据xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private List<SheetData> GetSheetDatas(ZipArchive zip)
        {
            var sheetDataList = new List<SheetData>();
            var sheetsEntry = zip.Entries.Where(x => x.FullName.StartsWith("xl/worksheets") && x.FullName.EndsWith(".xml") && !string.IsNullOrWhiteSpace(x.Name)).ToList();
            if (sheetsEntry == null)
                return sheetDataList;
            sheetsEntry.ForEach(sheetDataItem =>
            {
                var item = sheetDataItem.DeserializeSheet();
                item.Name = Path.GetFileNameWithoutExtension(sheetDataItem.Name);
                sheetDataList.Add(item);
            });
            return sheetDataList;
        }

        #endregion

        /// <summary>
        /// 保存为文件
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {

        }

        public void Dispose()
        {
            SharedStrings = null;
            Style = null;
            Sheets = null;
        }
    }
}
