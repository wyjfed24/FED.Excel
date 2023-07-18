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
        /// 行索引
        /// </summary>
        public int RowIndex { get; }
        
        /// <summary>
        /// 列号
        /// </summary>
        public string Column { get; }

        public ExcelSheetCell(int rowIndex, string column)
        {
            RowIndex = rowIndex;
            Column = column;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T GetValue<T>()
        {
            return (T)Value;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        public void SetValue<T>(T value)
        {
            Value = value;
        }
    }
}
