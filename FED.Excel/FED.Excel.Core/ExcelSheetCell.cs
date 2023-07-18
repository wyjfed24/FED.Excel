using FED.Excel.Core.Utility;

namespace FED.Excel.Core
{
    public class ExcelSheetCell
    {
        public object Value { get; set; }

        public int RowIndex { get; }
        public string Column { get; }

        public ExcelSheetCell(int rowIndex, string column)
        {
            RowIndex = rowIndex;
            Column = column;
        }

        public T GetValue<T>()
        {
            return (T)Value;
        }

        public void SetValue<T>(T value)
        {
            Value = value;
        }
    }
}
