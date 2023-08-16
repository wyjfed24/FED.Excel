// See https://aka.ms/new-console-template for more information
using FED.Excel.Core;
using FED.Excel.Core.Ext;
using FED.Excel.Core.Utility;
using FED.Excel.Test;

using System.Diagnostics;

//var sheet01 = wb.Sheets[0].ConvertTo<Sheet01>();
var i = 0;
//while (i < 100)
//{
    var w = new Stopwatch();
    w.Start();
    i++;
    var wb = new ExcelWorkbook<BigTest>("D:\\bigExcel.xlsx");
   // var sheet02 = wb.Sheets[0].ConvertTo<BigTest>();
    w.Stop();
    Console.WriteLine(w.ElapsedMilliseconds);

//}
Console.ReadLine();
//var sheet02 = wb.Sheets[1].ConvertTo<Sheet02>();

