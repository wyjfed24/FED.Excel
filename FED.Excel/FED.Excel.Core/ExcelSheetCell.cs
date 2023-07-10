namespace FED.Excel.Core
{
    public class ExcelSheetCell
    {
        public object Value { get; set; }

        public int RowIndex { get; }
        public int CellIndex { get; set; }
        public string Name { get { return ""; } }

        public ExcelSheetCell(int rowIndex, int cellIndex)
        {
            RowIndex = rowIndex;
            CellIndex = cellIndex;
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
