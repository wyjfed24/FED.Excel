using FED.Excel.Core.Utility;

namespace FED.Excel.Core
{
    public class ExcelSheetCell
    {
        /// <summary>
        /// 单元格值
        /// </summary>
        public object Value { get; set; }
        
        /// <summary>
        /// 列号
        /// </summary>
        public string Column { get; set; }

    }
}
