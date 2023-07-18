using System;
using System.Collections.Generic;
using System.Text;

namespace FED.Excel.Core.Utility
{
    public static class ExcelNumberConverter
    {
        private static string _baseNumberDic = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string ToExcelNumber(this int number)
        {
            var quotient = number % 26;
            var per = number / 26;
            var sb = new StringBuilder(_baseNumberDic[quotient].ToString());
            if (per != 0)
                sb.Insert(0, ToExcelNumber(per));
            return sb.ToString();
        }
    }
}
