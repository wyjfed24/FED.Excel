using FED.Excel.Core.ExcelXmlModel;
using FED.Excel.Core.Ext;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
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
            SharedStrings = GetShareStrings();
            Style = GetStyles();
        }

        #region Excel文件解析

        /// <summary>
        /// 获取共享字符串表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private ShareStringsTable GetShareStrings()
        {
            var sharedStringsEntry = _zip.Entries.Where(x => x.FullName == "xl/sharedStrings.xml").FirstOrDefault();
            if (sharedStringsEntry == null)
                return new ShareStringsTable();
            var sharedStringsTable = sharedStringsEntry.GetShareStrings();
            return sharedStringsTable;
        }

        /// <summary>
        /// 获取样式表xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        private StyleConfig GetStyles()
        {
            var styleEntry = _zip.Entries.Where(x => x.FullName == "xl/styles.xml").FirstOrDefault();
            if (styleEntry == null)
                return new StyleConfig();
            var style = styleEntry.GetStyleConfig();
            return style;
        }

        /// <summary>
        /// 获取Sheet数据xml
        /// </summary>
        /// <param name="zip"></param>
        /// <returns></returns>
        public List<T> GetSheetDatas<T>() where T : class, new()
        {
            var sheetDataList = new List<T>();
            var sheetsEntry = _zip.Entries.Where(x => x.FullName.StartsWith("xl/worksheets") && x.FullName.EndsWith(".xml") && !string.IsNullOrWhiteSpace(x.Name)).ToList();
            if (sheetsEntry == null)
                return sheetDataList;
            sheetsEntry.ForEach(sheetDataItem =>
            {
                var readHeader = false;
                List<(string Column, PropertyInfo Prop)> columnPropertyRelations = null;
                foreach (var rowCells in sheetDataItem.GetSheetRows(SharedStrings, Style))
                {
                    if (!readHeader)
                    {
                        columnPropertyRelations = rowCells.GetColumnPropertyRelations<T>();
                        readHeader = true;
                    }
                    else
                    {
                        var item = rowCells.ConvertTo<T>(columnPropertyRelations);
                        sheetDataList.Add(item);
                    }

                }
                //item.Name = Path.GetFileNameWithoutExtension(sheetDataItem.Name);
                //sheetDataList.Add(item);
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
