using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FED.Excel.Core.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string Name { get; set; }

        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }
    }
}
