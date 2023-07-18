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
            //获取表头属性映射集合
            var columnPropertyRelations = sheet.GetColumnPropertyRelations<T>();
            //跳过第一行转化数据到对象
            sheet.Rows.Skip(1).ToList().ForEach(x =>
            {
                var obj = x.ConvertTo<T>(sheet.Name, columnPropertyRelations);
                list.Add(obj);
            });
            return list;
        }

        /// <summary>
        /// 转换行到对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnPropertyRelations"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static T ConvertTo<T>(this ExcelSheetRow row, string sheetName, List<(string Column, PropertyInfo Prop)> columnPropertyRelations) where T : class, new()
        {
            var obj = new T();
            columnPropertyRelations.ForEach(x =>
            {
                var cell = row.Cells.FirstOrDefault(c => c.Column == x.Column);//由于excel源文件中未存储空单元格，造成每行单元格不一定索引对齐，所以通过列号来查询
                if (cell == null)
                    return;
                var cellValue = cell.Value;
                var propertyType = x.Prop.PropertyType;
                //处理可空类型
                var isNullabled = propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
                //获取到真实类型
                var realType = isNullabled ? propertyType.GetGenericArguments()[0] : propertyType;
                //枚举处理
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

        /// <summary>
        /// 获取表头和属性映射集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static List<(string Column, PropertyInfo Prop)> GetColumnPropertyRelations<T>(this ExcelWorksheet sheet)
        {
            //获取标记了特性的属性集合
            var type = typeof(T);
            var props = type.GetProperties();
            var excelProps = props.Where(x => x.CustomAttributes.Any(c => c.AttributeType == typeof(ExcelColumnAttribute)))
                                 .Select(x => (Column: string.Empty, Field: x.GetCustomAttribute<ExcelColumnAttribute>().Name, Prop: x))
                                 .ToList();
            //获取excel首行表头单元格
            var headCells = sheet.Rows.FirstOrDefault().Cells;
            //建立映射关系
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
