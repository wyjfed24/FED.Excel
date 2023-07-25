using System;
using System.Collections.Generic;
using System.Text;

namespace FED.Excel.Core.ExcelXmlModel
{
    /// <summary>
    /// 共享字符串表
    /// </summary>
    public class ShareStringsTable
    {
        public Dictionary<int, string> ShareStrings { get; set; } = new Dictionary<int, string>();

        public string this[int index] => ShareStrings[index];

        public void AddItem(int index, string content)
        {
            ShareStrings.Add(index, content);
        }
    }
}
