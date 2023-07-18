// See https://aka.ms/new-console-template for more information
using FED.Excel.Core;
using FED.Excel.Core.Ext;
using FED.Excel.Core.Utility;
using FED.Excel.Test;
Console.WriteLine("Hello, World!");
var wb = new ExcelWorkbook("D:\\test.xlsx");
var sheet01 = wb.Sheets[0].ConvertTo<Sheet01>();
var sheet02 = wb.Sheets[1].ConvertTo<Sheet02>();
Console.ReadLine();
