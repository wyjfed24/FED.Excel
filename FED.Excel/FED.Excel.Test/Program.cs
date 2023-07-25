// See https://aka.ms/new-console-template for more information
using FED.Excel.Core;
using FED.Excel.Core.Ext;
using FED.Excel.Core.Utility;
using FED.Excel.Test;

using System.Diagnostics;

var w = new Stopwatch();
w.Start();
var wb = new ExcelWorkbook("D:\\test.xlsx");
//var sheet01 = wb.Sheets[0].ConvertTo<Sheet01>();
var sheet02 = wb.Sheets[1].ConvertTo<Sheet02>();
w.Stop();
Console.WriteLine(w.ElapsedMilliseconds);
Console.ReadLine();
//var sheet02 = wb.Sheets[1].ConvertTo<Sheet02>();

