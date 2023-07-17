using FED.Excel.Core.Attributes;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FED.Excel.Test
{
    public class Test001
    {
        [ExcelColumn("序号")]
        public int No { get; set; }

        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("日期")]
        public DateTime Date { get; set; }

        public int Sort { get; set; }

        [ExcelColumn("备注")]
        public string Remark { get; set; }
    }
}
