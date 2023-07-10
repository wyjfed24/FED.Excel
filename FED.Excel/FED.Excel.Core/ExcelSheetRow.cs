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

        public ExcelSheetCell AppendCell()
        {
            var cell = new ExcelSheetCell(RowIndex, Cells.Count);
            Cells.Add(cell);
            return cell;
        }

        public ExcelSheetCell InsertCell(int index)
        {
            var maxIndex = Cells.Count - 1;
            if (index > maxIndex)
                return AppendCell();
            var cell = new ExcelSheetCell(RowIndex, index);
            Cells.Where(x => x.CellIndex >= index).ToList().ForEach(x => x.CellIndex += 1);
            Cells.Insert(index, cell);
            return cell;
        }
    }
}
