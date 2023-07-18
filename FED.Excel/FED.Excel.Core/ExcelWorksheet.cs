using System.Collections.Generic;
using System.Linq;

namespace FED.Excel.Core
{
    public class ExcelWorksheet
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public List<ExcelSheetRow> Rows { get; set; } = new List<ExcelSheetRow>();

        public ExcelWorksheet(int index, string name) {
            Index = index;
            Name = name;
        }

        public ExcelSheetRow AppendRow()
        {
            var row = new ExcelSheetRow(Rows.Count);
            Rows.Add(row);
            return row;
        }

        public ExcelSheetRow InsertRow(int index)
        {
            var maxIndex = Rows.Count - 1;
            if (index > maxIndex)
                return AppendRow();
            var row = new ExcelSheetRow(index);
            Rows.Where(x => x.RowIndex >= index).ToList().ForEach(x => x.RowIndex += 1);
            Rows.Insert(index, row);
            return row;
        }
    }
}
