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
                var obj = x.ConvertTo<T>(sheet.Name, columns);
                list.Add(obj);
            });
            return list;
        }

        private static T ConvertTo<T>(this ExcelSheetRow row, string sheetName, List<(string Column, PropertyInfo Prop)> columns) where T : class, new()
        {
            var obj = new T();
            columns.ForEach(x =>
            {
                var cell = row.Cells.FirstOrDefault(c => c.Column == x.Column);
                if (cell == null)
                    return;
                var cellValue = cell.Value;
                var isNullabled = x.Prop.PropertyType.IsGenericType && x.Prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
                var propertyType = x.Prop.PropertyType;
                var realType = isNullabled ? propertyType.GetGenericArguments()[0] : propertyType;
                if (realType.IsEnum)
                {
                    try
                    {
                        var enumValue = System.Enum.Parse(realType, cellValue.ToString());
                        x.Prop.SetValue(obj, enumValue);
                    }
                    catch
                    {
                        throw new Exception($"Sheet：{sheetName}，行号：{row.RowIndex + 1}，列号：{cell.Column}，值：{cellValue} 不能转化为[{realType.Name}]枚举");
                    }
                }
                else
                {
                    var value = Convert.ChangeType(cellValue, realType);
                    x.Prop.SetValue(obj, value);
                }
            });
            return obj;
        }

        private static List<(string Column, PropertyInfo Prop)> GetColumns<T>(this ExcelWorksheet sheet)
        {
            var type = typeof(T);
            var props = type.GetProperties();
            var excelProps = props.Where(x => x.CustomAttributes.Any(c => c.AttributeType == typeof(ExcelColumnAttribute)))
                                 .Select(x => (Column: string.Empty, Field: x.GetCustomAttribute<ExcelColumnAttribute>().Name, Prop: x))
                                 .ToList();
            var headCells = sheet.Rows.FirstOrDefault().Cells;
            var columns = new List<(string Column, PropertyInfo Prop)>();
            for (int i = 0; i < excelProps.Count; i++)
            {
                var item = excelProps[i];
                var cell = headCells.FirstOrDefault(c => c.GetValue<string>() == item.Field);
                if (cell == null)
                    continue;
                columns.Add((cell.Column, item.Prop));
            }
            return columns;
        }
    }
}
