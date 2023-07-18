using System.Collections.Generic;
using System.Linq;

namespace FED.Excel.Core
{
    public class ExcelSheetRow
    {
        public int RowIndex { get; set; }
        public List<ExcelSheetCell> Cells { get; set; } = new List<ExcelSheetCell>();

        public ExcelSheetRow(int rowIndex)
        {
            RowIndex = rowIndex;
        }

        public ExcelSheetCell CreateCell(string column)
        {
            var cell = new ExcelSheetCell(RowIndex, column);
            Cells.Add(cell);
            return cell;
        }
    }
}
