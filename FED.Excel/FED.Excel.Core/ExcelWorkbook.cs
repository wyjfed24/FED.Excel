using FED.Excel.Core.Ext;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace FED.Excel.Core
{
    public class ExcelWorkbook : IDisposable
    {
        public List<ExcelWorksheet> Sheets { get; set; } = new List<ExcelWorksheet>();

        public ExcelWorkbook() { }

        public ExcelWorkbook(Stream stream)
        {
            BuildWorkbook(stream);
        }

        public ExcelWorkbook(string filePath)
        {
            BuildWorkbook(filePath);
        }

        private void BuildWorkbook(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                BuildWorkbook(stream);
            }
        }

        private void BuildWorkbook(Stream stream)
        {
            using (var package = new ExcelPackage(stream))
            {
                //转换对象
                BuildWorkbook(package);
            }
        }

        private void BuildWorkbook(ExcelPackage package)
        {
            var pgUnNullSheets = package.Sheets.Where(x => x.SheetData.Rows != null).ToList();
            foreach (var pgSheet in pgUnNullSheets)
            {
                var sheet = AppendSheet(pgSheet.Name);
                foreach (var pgRow in pgSheet.SheetData.Rows)
                {
                    var row = sheet.AppendRow();
                    foreach (var pgCell in pgRow.Cells)
                    {
                        var cell = row.CreateCell(pgCell.CellNumber.Replace(pgRow.RowNumber.ToString(), string.Empty));
                        if (pgCell.Value == null)
                            continue;
                        if (pgCell.CellType == "s")//字符串
                        {
                            string value;
                            try//先转换为索引查询公共字符串表
                            {
                                var index = Convert.ToInt32(pgCell.Value);
                                value = package.SharedStringsTable[index];
                            }
                            catch//失败则为原始值
                            {
                                value = pgCell.Value;
                            }
                            cell.SetValue(value);
                        }
                        else
                        {
                            //判断是日期还是数字
                            if (package.Style.IsDate(pgCell.StyleId))
                            {
                                var sourceValue = Convert.ToDouble(pgCell.Value);
                                var value = DateTime.FromOADate(sourceValue);
                                cell.SetValue(value);
                            }
                            else
                            {
                                cell.SetValue(pgCell.Value);
                            }
                        }
                    }
                }
            }
        }


        public ExcelWorksheet AppendSheet(string name)
        {
            var sheet = new ExcelWorksheet(Sheets.Count, name);
            Sheets.Add(sheet);
            return sheet;
        }

        public ExcelWorksheet InsertSheet(int index, string name)
        {
            var maxIndex = Sheets.Count - 1;
            if (index > maxIndex)
                return AppendSheet(name);
            var sheet = new ExcelWorksheet(index, name);
            Sheets.Where(x => x.Index >= index).ToList().ForEach(x => x.Index += 1);
            Sheets.Insert(index, sheet);
            return sheet;
        }

        /// <summary>
        /// 保存为文件
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            //转package再调用SaveAs()
        }

        public void Dispose()
        {
            Sheets = null;
        }
    }
}
