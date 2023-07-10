using FED.Excel.Core.ExcelXmlModel;

using System.Linq;

namespace FED.Excel.Core.Ext
{
    public static class ZipXmlValidateExt
    {
        public static bool IsDate(this StyleXmlModel style, int styleId)
        {
            var cellXfs = style.CellXfInfo.CellXfs[styleId];
            if (cellXfs.NumFmtId == 14)
                return true;
            var numberFormat = style.NumberFormatInfo.NumberFormats.Where(x => x.NumFmtId == cellXfs.NumFmtId).FirstOrDefault();
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
