using FED.Excel.Core.ExcelXmlModel;
using FED.Excel.Core.Ext;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace FED.Excel.Core
{
    internal class ExcelPackage : IDisposable
    {
        /// <summary>
        /// 共享字符串表
        /// </summary>
        internal SharedStringsTableXmlModel SharedStringsTable { get; set; }

        /// <summary>
        /// 样式表
        /// </summary>
        internal StyleXmlModel Style { get; set; }

        /// <summary>
        /// Sheet原始数据集合
        /// </summary>
        internal List<SheetXmlModel> Sheets { get; set; }

        internal ExcelPackage(Stream stream)
        {
            var zip = new ZipArchive(stream);
            SharedStringsTable = GetShareStrings(zip);
            Style = GetStyles(zip);
            Sheets = GetSheetDatas(zip);
        }

        #region Excel文件解析

        /// <summary>
        /// 反序列化共享字符串表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private SharedStringsTableXmlModel GetShareStrings(ZipArchive zip)
        {
            var sharedStringsEntry = zip.Entries.Where(x => x.FullName == "xl/sharedStrings.xml").FirstOrDefault();
            if (sharedStringsEntry == null)
                return new SharedStringsTableXmlModel();
            var sharedStringsTable = sharedStringsEntry.Deserialize<SharedStringsTableXmlModel>();
            return sharedStringsTable;
        }

        /// <summary>
        /// 反序列化样式表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private StyleXmlModel GetStyles(ZipArchive zip)
        {
            var styleEntry = zip.Entries.Where(x => x.FullName == "xl/styles.xml").FirstOrDefault();
            if (styleEntry == null)
                return new StyleXmlModel();
            var style = styleEntry.Deserialize<StyleXmlModel>();
            return style;
        }

        /// <summary>
        /// 反序列化Sheet数据xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private List<SheetXmlModel> GetSheetDatas(ZipArchive zip)
        {
            var sheetsEntry = zip.Entries.Where(x => x.FullName.StartsWith("xl/worksheets") && x.FullName.EndsWith(".xml") && !string.IsNullOrWhiteSpace(x.Name)).ToList();
            if (sheetsEntry == null)
                return new List<SheetXmlModel>();
            var sheetDataList = new List<SheetXmlModel>();
            sheetsEntry.ForEach(sheetDataItem => sheetDataList.Add(sheetDataItem.Deserialize<SheetXmlModel>()));
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
            SharedStringsTable = null;
            Style = null;
            Sheets = null;
        }
    }
}
