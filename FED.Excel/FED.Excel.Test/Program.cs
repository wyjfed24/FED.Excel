// See https://aka.ms/new-console-template for more information
using FED.Excel.Core;
using FED.Excel.Core.Ext;
using FED.Excel.Test;

Console.WriteLine("Hello, World!");
var wb = new ExcelWorkbook("D:\\1.xlsx");
var test = wb.Sheets.FirstOrDefault().ConvertTo<Test001>(false);
Console.ReadLine();