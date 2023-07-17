using FED.Excel.Core.Attributes;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Schema;

namespace FED.Excel.Core.Ext
{
    public static class ExcelConverterExt
    {
        public static List<T> ConvertTo<T>(this ExcelWorksheet sheet) where T : class, new()
        {
            List<T> list = new List<T>();
            var columns = sheet.GetColumns<T>();
            sheet.Rows.Skip(1).ToList().ForEach(x =>
            {
                var obj = x.ConvertTo<T>(columns);
                list.Add(obj);
            });
            return list;
        }

        private static T ConvertTo<T>(this ExcelSheetRow row, List<(int CellIndex, PropertyInfo Prop)> columns) where T : class, new()
        {
            var obj = new T();
            columns.ForEach(x =>
            {
                var cellValue = row.Cells[x.CellIndex].Value;
                var value = Convert.ChangeType(cellValue, x.Prop.PropertyType);
                x.Prop.SetValue(obj, value);
            });
            return obj;
        }

        private static List<(int CellIndex, PropertyInfo Prop)> GetColumns<T>(this ExcelWorksheet sheet)
        {
            var type = typeof(T);
            var props = type.GetProperties();
            var excelProps = props.Where(x => x.CustomAttributes.Any(c => c.AttributeType == typeof(ExcelColumnAttribute)))
                                 .Select(x => (CellIndex: -1, Field: x.GetCustomAttribute<ExcelColumnAttribute>().Name, Prop: x))
                                 .ToList();
            var headCells = sheet.Rows.FirstOrDefault().Cells;
            var columns = new List<(int CellIndex, PropertyInfo Prop)>();
            for (int i = 0; i < excelProps.Count; i++)
            {
                var item = excelProps[i];
                var cell = headCells.FirstOrDefault(c => c.GetValue<string>() == item.Field);
                if (cell == null)
                    continue;
                columns.Add((cell.CellIndex, item.Prop));
            }
            return columns;
        }
    }
}
