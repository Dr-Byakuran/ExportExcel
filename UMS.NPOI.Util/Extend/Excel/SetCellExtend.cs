using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// 设置值扩展
    /// </summary>
    public static class SetCellExtend
    {
        private static IWorkbook workBook = null;
        private static ISheet sheet = null;
        private static ICellStyle cellStyle = null;

        /// <summary>
        /// 设置值：数值类型（自定义、货币、会计专用）
        /// <para>先设置样式，再调用此赋值</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBook"></param>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <param name="dot"></param>
        public static void SetCellValue(this ICell cell, object value, ExcelNumberType numberType, ExcelCurrencyType currencyType = ExcelCurrencyType.人民币, int dot = 2)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            string strDataFormat = string.Empty;
            string dt = ExcelExtend.GetDot(dot);
            string ct = string.Empty;
            // 设置货币类型
            switch (currencyType)
            {
                case ExcelCurrencyType.人民币:
                    ct = "￥";
                    break;
                case ExcelCurrencyType.美元:
                    ct = "$";
                    break;
                case ExcelCurrencyType.欧元:
                    ct = "€";
                    break;
                case ExcelCurrencyType.英镑:
                    ct = "£";
                    break;
            }
            IDataFormat sdf = workBook.CreateDataFormat();
            // 获取类型
            switch (numberType)
            {
                case ExcelNumberType.会计专用1:
                    strDataFormat = " _ * #,##0_ ;_ * -#,##0_ ;_ * ' - '_ ;_ @_ ";
                    break;
                case ExcelNumberType.会计专用2:
                    strDataFormat = "_ * #,##0{0}_ ;_ * -#,##0{0}_ ;_ * ' - '??_ ;_ @_ ";
                    strDataFormat = string.Format(strDataFormat, dt);
                    break;
                case ExcelNumberType.会计专用3:
                    strDataFormat = "_ {0}* #,##0_ ;_ {0}* -#,##0_ ;_ {0}* ' - '_ ;_ @_ ";
                    strDataFormat = string.Format(strDataFormat, ct);
                    break;
                case ExcelNumberType.会计专用4:
                    strDataFormat = "_ {0}* #,##0{1}_ ;_ {0}* -#,##0{1}_ ;_ {0}* ' - '??_ ;_ @_ ";
                    strDataFormat = string.Format(strDataFormat, ct, dt);
                    break;
                case ExcelNumberType.自定义1:
                    strDataFormat = "0";
                    break;
                case ExcelNumberType.自定义2:
                    strDataFormat = "0.00";
                    break;
                case ExcelNumberType.自定义3:
                    strDataFormat = "#,##0;-#,##0";
                    break;
                case ExcelNumberType.自定义4:
                    strDataFormat = "#,##0{0};-#,##0{0}";
                    strDataFormat = string.Format(strDataFormat, dt);
                    break;
                case ExcelNumberType.货币1:
                    strDataFormat = "#,##0";
                    break;
                case ExcelNumberType.货币2:
                    strDataFormat = "#,##0{0}";
                    strDataFormat = string.Format(strDataFormat, dt);
                    break;
                case ExcelNumberType.货币3:
                    strDataFormat = "#,##0;[Red]-#,##0";
                    break;
                case ExcelNumberType.货币4:
                    strDataFormat = "#,##0{0};[Red]-#,##0{0}";
                    strDataFormat = string.Format(strDataFormat, dt);
                    break;
                case ExcelNumberType.货币5:
                    strDataFormat = "{0}#,##0;{0}-#,##0";
                    strDataFormat = string.Format(strDataFormat, ct);
                    break;
                case ExcelNumberType.货币6:
                    strDataFormat = "{0}#,##0;[Red]{0}-#,##0";
                    strDataFormat = string.Format(strDataFormat, ct);
                    break;
                case ExcelNumberType.货币7:
                    strDataFormat = "{0}#,##0{1};{0}-#,##0{1}";
                    strDataFormat = string.Format(strDataFormat, ct, dt);
                    break;
                case ExcelNumberType.货币8:
                    strDataFormat = "{0}#,##0{1};[Red]{0}-#,##0{1}";
                    strDataFormat = string.Format(strDataFormat, ct, dt);
                    break;
                case ExcelNumberType.货币9:
                    strDataFormat = "{0}#,##0_);({0}#,##0)";
                    strDataFormat = string.Format(strDataFormat, ct);
                    break;
                case ExcelNumberType.货币10:
                    strDataFormat = "{0}#,##0_);[Red]({0}#,##0)";
                    strDataFormat = string.Format(strDataFormat, ct);
                    break;
                case ExcelNumberType.货币11:
                    strDataFormat = "{0}#,##0{1}_);({0}#,##0{1})";
                    strDataFormat = string.Format(strDataFormat, ct, dt);
                    break;
                case ExcelNumberType.货币12:
                    strDataFormat = "{0}#,##0{1}_);[Red]({0}#,##0{1})";
                    strDataFormat = string.Format(strDataFormat, ct, dt);
                    break;
                default:
                    break;
            }
            double dbValue = 0;
            double.TryParse(value.ToString(), out dbValue);
            // 设置值
            cell.SetCellValue(dbValue);
            // 设置样式
            cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置值：科学计数法
        /// <para>先设置样式，再调用此赋值</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBook"></param>
        /// <param name="value"></param>
        /// <param name="countType"></param>
        /// <param name="cellStyle"></param>
        public static void SetCellValue(this ICell cell, object value, ExcelCountType countType, int dot = 2)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            string dt = ExcelExtend.GetDot(dot);
            string strDataFormat = string.Empty;
            switch (countType)
            {
                case ExcelCountType.科学计数1:
                    strDataFormat = "0{0}E+00";
                    if (dot == 0)
                        dt = ".";
                    strDataFormat = string.Format(strDataFormat, dt);
                    break;
                case ExcelCountType.自定义1:
                    strDataFormat = "##0.0E+0";
                    break;
            }
            double dbValue = 0;
            double.TryParse(value.ToString(), out dbValue);
            cell.SetCellValue(dbValue);
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置类型：日期、时间
        /// <para>先设置样式，再调用此赋值</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBook"></param>
        /// <param name="value"></param>
        /// <param name="dateType"></param>
        /// <param name="cellStyle"></param>
        public static void SetCellValue(this ICell cell, object value, ExcelDateType dateType)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            string strDataFormat = string.Empty;
            // 获取类型
            switch (dateType)
            {
                case ExcelDateType.日期1:
                    strDataFormat = "yyyy/m/d";
                    break;
                case ExcelDateType.日期2:
                    strDataFormat = "[$-F800]dddd, mmmm dd, yyyy";
                    break;
                case ExcelDateType.日期3:
                    strDataFormat = "[DBNum1][$-804]yyyy年m月d日;@";
                    break;
                case ExcelDateType.日期4:
                    strDataFormat = "[DBNum1][$-804]yyyy年m月;@";
                    break;
                case ExcelDateType.日期5:
                    strDataFormat = "[DBNum1][$-804]m月d日;@";
                    break;
                case ExcelDateType.日期6:
                    strDataFormat = "[$-804]aaaa;@";
                    break;
                case ExcelDateType.日期7:
                    strDataFormat = "[$-804]aaa;@";
                    break;
                case ExcelDateType.时间1:
                    strDataFormat = "[DBNum1][$-804]h时mm分;@";
                    break;
                case ExcelDateType.时间2:
                    strDataFormat = "[DBNum1][$-804]上午/下午h时mm分;@";
                    break;
                case ExcelDateType.自定义1:
                    strDataFormat = "yyyy年m月";
                    break;
                case ExcelDateType.自定义2:
                    strDataFormat = "m月d日";
                    break;
                case ExcelDateType.自定义3:
                    strDataFormat = "yyyy年m月d日";
                    break;
                case ExcelDateType.自定义4:
                    strDataFormat = "m/d/yy";
                    break;
                case ExcelDateType.自定义5:
                    strDataFormat = "d-mmm-yy";
                    break;
                case ExcelDateType.自定义6:
                    strDataFormat = "d-mmm";
                    break;
                case ExcelDateType.自定义7:
                    strDataFormat = "mmm-yy";
                    break;
                case ExcelDateType.自定义8:
                    strDataFormat = "h:mm AM/PM";
                    break;
                case ExcelDateType.自定义9:
                    strDataFormat = "h:mm:ss AM/PM";
                    break;
                case ExcelDateType.自定义10:
                    strDataFormat = "h:mm";
                    break;
                case ExcelDateType.自定义11:
                    strDataFormat = "h:mm:ss";
                    break;
                case ExcelDateType.自定义12:
                    strDataFormat = "h时mm分";
                    break;
                case ExcelDateType.自定义13:
                    strDataFormat = "h时mm分ss秒";
                    break;
                case ExcelDateType.自定义14:
                    strDataFormat = "上午/下午h时mm分";
                    break;
                case ExcelDateType.自定义15:
                    strDataFormat = "上午/下午h时mm分ss秒";
                    break;
                case ExcelDateType.自定义16:
                    strDataFormat = "yyyy/m/d h:mm";
                    break;
                case ExcelDateType.自定义17:
                    strDataFormat = "mm:ss";
                    break;
                case ExcelDateType.自定义18:
                    strDataFormat = "mm:ss.0";
                    break;
            }
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置值：分数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBook"></param>
        /// <param name="value"></param>
        /// <param name="facType"></param>
        /// <param name="cellStyle"></param>
        public static void SetCellValue(this ICell cell, object value, ExcelFractionType facType)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            string strDataFormat = string.Empty;
            switch (facType)
            {
                case ExcelFractionType.分数1:
                    strDataFormat = "# ?/?";
                    break;
                case ExcelFractionType.分数2:
                    strDataFormat = "# ??/??";
                    break;
                case ExcelFractionType.分数3:
                    strDataFormat = "# ???/???";
                    break;
                case ExcelFractionType.分数4:
                    strDataFormat = "# ?/2";
                    break;
                case ExcelFractionType.分数5:
                    strDataFormat = "# ?/4";
                    break;
                case ExcelFractionType.分数6:
                    strDataFormat = "# ?/8";
                    break;
                case ExcelFractionType.分数7:
                    strDataFormat = "# ?/16";
                    break;
                case ExcelFractionType.分数8:
                    strDataFormat = "# ?/10";
                    break;
                case ExcelFractionType.分数9:
                    strDataFormat = "# ??/100";
                    break;
            }
            DateTime dtDate = DateTime.Parse(value.ToString());
            cell.SetCellValue(dtDate);
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置值：特殊类型 数字转中文
        /// <para>先设置样式，再调用此赋值</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBook"></param>
        /// <param name="value"></param>
        /// <param name="specType"></param>
        /// <param name="cellStyle"></param>
        public static void SetCellValue(this ICell cell, object value, ExcelSpecialType specType)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            string strDataFormat = string.Empty;
            switch (specType)
            {
                case ExcelSpecialType.特殊1:
                    strDataFormat = "000000";
                    break;
                case ExcelSpecialType.特殊2:
                    strDataFormat = "[DBNum1][$-804]G/通用格式";
                    break;
                case ExcelSpecialType.特殊3:
                    strDataFormat = "[DBNum2][$-804]G/通用格式";
                    break;
            }
            double dbValue = 0;
            double.TryParse(value.ToString(), out dbValue);
            cell.SetCellValue(dbValue);
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
            cell.CellStyle = cellStyle;

        }

        /// <summary>
        /// 设置值：自定义 DataFormat
        /// <para>先设置样式，再调用此赋值</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataFormat"></param>
        private static void SetCellValue(this ICell cell, object value, string dataFormat)
        {
            if (cell == null)
                return;
            cell.DealParam();

            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat(dataFormat);
            cell.SetCellValue(value.ToString());
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 处理参数
        /// <para>参数：IWorkBook, ISheet, ICellStyle</para>
        /// </summary>
        /// <param name="cell"></param>
        private static void DealParam(this ICell cell)
        {
            cellStyle = cell.CellStyle;
            sheet = cell.Sheet;
            workBook = sheet.Workbook;
            if (cellStyle == null)
                cellStyle = workBook.CreateCellStyle();
        }
    }
}
