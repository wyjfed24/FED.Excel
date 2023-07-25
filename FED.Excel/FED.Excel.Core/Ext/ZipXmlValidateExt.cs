using FED.Excel.Core.ExcelXmlModel;

using System.Linq;

namespace FED.Excel.Core.Ext
{
    public static class ZipXmlValidateExt
    {
        public static bool IsDate(this StyleConfig style, int styleId)
        {
            var cellXfs = style.CellXfs[styleId];
            if (cellXfs.ApplyNumFmtId == 14)//日期默认样式Id为14
                return true;
            //自定义日期需处理格式化规则
            var numberFormat = style.NumberFormats.Where(x => x.NumFmtId == cellXfs.ApplyNumFmtId).FirstOrDefault();
            if (numberFormat == null)
                return false;
            var formatCode = numberFormat.FormatCode.ToLower().Replace("red","");//排除掉Red标签
            return formatCode.Contains("yy") ||
                 formatCode.Contains("m") ||
                 formatCode.Contains("d") ||
                 formatCode.Contains("h") ||
                 formatCode.Contains("s") ||
                 formatCode.Contains("aaa") ||
                 formatCode.Contains("am") ||
                 formatCode.Contains("ap");
        }
    }
}
