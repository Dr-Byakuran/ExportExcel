using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Web;
using System.Text;
using System.Linq;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// Npoi Excel 扩展帮助
    /// </summary>
    public static class NpoiExcelExtend
    {
  //      private static readonly short defaultColorIndexed = 9;
  //      private static IWorkbook workBook = null;
  //      private static ISheet sheet = null;
  //      private static ICellStyle cellStyle = null;
  //      private static List<Tuple<int, string>> originalRGBs = new List<Tuple<int, string>>();

  //      /// <summary>
  //      /// 写入Http
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="fileName">文件名称</param>
  //      public static void HttpWrite(this IWorkbook workBook, string fileName)
  //      {
  //          using (MemoryStream ms = new MemoryStream())
  //          {
  //              workBook.Write(ms);
  //              HttpContext.Current.Response.Clear();
  //              HttpContext.Current.Response.ClearHeaders();
  //              HttpContext.Current.Response.Buffer = true;
  //              if (!Contains(HttpContext.Current.Request.UserAgent, "firefox", true) &&
  //              !Contains(HttpContext.Current.Request.UserAgent, "chrome", true))
  //                  fileName = UrlEncode(fileName, Encoding.UTF8, false);
  //              fileName = fileName.Replace("\"", "");
  //              HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + fileName);
  //              // 加入ContentType 防止火狐浏览器导出时直接导出Html，让其默认Excel导出
  //              HttpContext.Current.Response.ContentType = "application/ms-excel";
  //              HttpContext.Current.Response.BinaryWrite(ms.ToArray());
  //              HttpContext.Current.Response.Flush();
  //              HttpContext.Current.Response.End();
  //              ms.Close();
  //          }
  //      }

  //      #region 扩展单元格赋值

  //      /// <summary>
  //      /// 设置值：数值类型（自定义、货币、会计专用）
  //      /// <para>先设置样式，再调用此赋值</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="workBook"></param>
  //      /// <param name="value"></param>
  //      /// <param name="cellStyle"></param>
  //      /// <param name="dot"></param>
  //      public static void SetCellValue(this ICell cell,  object value, ExcelNumberType numberType, ExcelCurrencyType currencyType = ExcelCurrencyType.人民币 , int dot = 2)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();

  //          if (value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          string strDataFormat = string.Empty;
  //          string dt = GetDot(dot);
  //          string ct = string.Empty;
  //          // 设置货币类型
  //          switch (currencyType)
  //          {
  //              case ExcelCurrencyType.人民币:
  //                  ct = "￥";
  //                  break;
  //              case ExcelCurrencyType.美元:
  //                  ct = "$";
  //                  break;
  //              case ExcelCurrencyType.欧元:
  //                  ct = "€";
  //                  break;
  //              case ExcelCurrencyType.英镑:
  //                  ct = "£";
  //                  break;
  //          }
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          // 获取类型
  //          switch (numberType)
  //          {
  //              case ExcelNumberType.会计专用1:
  //                  strDataFormat = " _ * #,##0_ ;_ * -#,##0_ ;_ * ' - '_ ;_ @_ ";
  //                  break;
  //              case ExcelNumberType.会计专用2:
  //                  strDataFormat = "_ * #,##0{0}_ ;_ * -#,##0{0}_ ;_ * ' - '??_ ;_ @_ ";
  //                  strDataFormat = string.Format(strDataFormat, dt);
  //                  break;
  //              case ExcelNumberType.会计专用3:
  //                  strDataFormat = "_ {0}* #,##0_ ;_ {0}* -#,##0_ ;_ {0}* ' - '_ ;_ @_ ";
  //                  strDataFormat = string.Format(strDataFormat, ct);
  //                  break;
  //              case ExcelNumberType.会计专用4:
  //                  strDataFormat = "_ {0}* #,##0{1}_ ;_ {0}* -#,##0{1}_ ;_ {0}* ' - '??_ ;_ @_ ";
  //                  strDataFormat = string.Format(strDataFormat, ct, dt);
  //                  break;
  //              case ExcelNumberType.自定义1:
  //                  strDataFormat = "0";
  //                  break;
  //              case ExcelNumberType.自定义2:
  //                  strDataFormat = "0.00";
  //                  break;
  //              case ExcelNumberType.自定义3:
  //                  strDataFormat = "#,##0;-#,##0";
  //                  break;
  //              case ExcelNumberType.自定义4:
  //                  strDataFormat = "#,##0{0};-#,##0{0}";
  //                  strDataFormat = string.Format(strDataFormat, dt);
  //                  break;
  //              case ExcelNumberType.货币1:
  //                  strDataFormat = "#,##0";
  //                  break;
  //              case ExcelNumberType.货币2:
  //                  strDataFormat = "#,##0{0}";
  //                  strDataFormat = string.Format(strDataFormat, dt);
  //                  break;
  //              case ExcelNumberType.货币3:
  //                  strDataFormat = "#,##0;[Red]-#,##0";
  //                  break;
  //              case ExcelNumberType.货币4:
  //                  strDataFormat = "#,##0{0};[Red]-#,##0{0}";
  //                  strDataFormat = string.Format(strDataFormat, dt);
  //                  break;
  //              case ExcelNumberType.货币5:
  //                  strDataFormat = "{0}#,##0;{0}-#,##0";
  //                  strDataFormat = string.Format(strDataFormat, ct);
  //                  break;
  //              case ExcelNumberType.货币6:
  //                  strDataFormat = "{0}#,##0;[Red]{0}-#,##0";
  //                  strDataFormat = string.Format(strDataFormat, ct);
  //                  break;
  //              case ExcelNumberType.货币7:
  //                  strDataFormat = "{0}#,##0{1};{0}-#,##0{1}";
  //                  strDataFormat = string.Format(strDataFormat, ct, dt);
  //                  break;
  //              case ExcelNumberType.货币8:
  //                  strDataFormat = "{0}#,##0{1};[Red]{0}-#,##0{1}";
  //                  strDataFormat = string.Format(strDataFormat, ct, dt);
  //                  break;
  //              case ExcelNumberType.货币9:
  //                  strDataFormat = "{0}#,##0_);({0}#,##0)";
  //                  strDataFormat = string.Format(strDataFormat, ct);
  //                  break;
  //              case ExcelNumberType.货币10:
  //                  strDataFormat = "{0}#,##0_);[Red]({0}#,##0)";
  //                  strDataFormat = string.Format(strDataFormat, ct);
  //                  break;
  //              case ExcelNumberType.货币11:
  //                  strDataFormat = "{0}#,##0{1}_);({0}#,##0{1})";
  //                  strDataFormat = string.Format(strDataFormat, ct, dt);
  //                  break;
  //              case ExcelNumberType.货币12:
  //                  strDataFormat = "{0}#,##0{1}_);[Red]({0}#,##0{1})";
  //                  strDataFormat = string.Format(strDataFormat, ct, dt);
  //                  break;
  //              default:
  //                  break;
  //          }
  //          double dbValue = 0;
  //          double.TryParse(value.ToString(), out dbValue);
  //          // 设置值
  //          cell.SetCellValue(dbValue);
  //          // 设置样式
  //          cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置值：科学计数法
  //      /// <para>先设置样式，再调用此赋值</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="workBook"></param>
  //      /// <param name="value"></param>
  //      /// <param name="countType"></param>
  //      /// <param name="cellStyle"></param>
  //      public static void SetCellValue(this ICell cell,  object value, ExcelCountType countType , int dot = 2)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();

  //          if (value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          string dt = GetDot(dot);
  //          string strDataFormat = string.Empty;
  //          switch (countType)
  //          {
  //              case ExcelCountType.科学计数1:
  //                  strDataFormat = "0{0}E+00";
  //                  if (dot == 0)
  //                      dt = ".";
  //                  strDataFormat = string.Format(strDataFormat, dt);
  //                  break;
  //              case ExcelCountType.自定义1:
  //                  strDataFormat = "##0.0E+0";
  //                  break;
  //          }
  //          double dbValue = 0;
  //          double.TryParse(value.ToString(), out dbValue);
  //          cell.SetCellValue(dbValue);
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置类型：日期、时间
  //      /// <para>先设置样式，再调用此赋值</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="workBook"></param>
  //      /// <param name="value"></param>
  //      /// <param name="dateType"></param>
  //      /// <param name="cellStyle"></param>
  //      public static void SetCellValue(this ICell cell,  object value, ExcelDateType dateType )
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();

  //          if (value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          string strDataFormat = string.Empty;
  //          // 获取类型
  //          switch (dateType)
  //          {
  //              case ExcelDateType.日期1:
  //                  strDataFormat = "yyyy/m/d";
  //                  break;
  //              case ExcelDateType.日期2:
  //                  strDataFormat = "[$-F800]dddd, mmmm dd, yyyy";
  //                  break;
  //              case ExcelDateType.日期3:
  //                  strDataFormat = "[DBNum1][$-804]yyyy年m月d日;@";
  //                  break;
  //              case ExcelDateType.日期4:
  //                  strDataFormat = "[DBNum1][$-804]yyyy年m月;@";
  //                  break;
  //              case ExcelDateType.日期5:
  //                  strDataFormat = "[DBNum1][$-804]m月d日;@";
  //                  break;
  //              case ExcelDateType.日期6:
  //                  strDataFormat = "[$-804]aaaa;@";
  //                  break;
  //              case ExcelDateType.日期7:
  //                  strDataFormat = "[$-804]aaa;@";
  //                  break;
  //              case ExcelDateType.时间1:
  //                  strDataFormat = "[DBNum1][$-804]h时mm分;@";
  //                  break;
  //              case ExcelDateType.时间2:
  //                  strDataFormat = "[DBNum1][$-804]上午/下午h时mm分;@";
  //                  break;
  //              case ExcelDateType.自定义1:
  //                  strDataFormat = "yyyy年m月";
  //                  break;
  //              case ExcelDateType.自定义2:
  //                  strDataFormat = "m月d日";
  //                  break;
  //              case ExcelDateType.自定义3:
  //                  strDataFormat = "yyyy年m月d日";
  //                  break;
  //              case ExcelDateType.自定义4:
  //                  strDataFormat = "m/d/yy";
  //                  break;
  //              case ExcelDateType.自定义5:
  //                  strDataFormat = "d-mmm-yy";
  //                  break;
  //              case ExcelDateType.自定义6:
  //                  strDataFormat = "d-mmm";
  //                  break;
  //              case ExcelDateType.自定义7:
  //                  strDataFormat = "mmm-yy";
  //                  break;
  //              case ExcelDateType.自定义8:
  //                  strDataFormat = "h:mm AM/PM";
  //                  break;
  //              case ExcelDateType.自定义9:
  //                  strDataFormat = "h:mm:ss AM/PM";
  //                  break;
  //              case ExcelDateType.自定义10:
  //                  strDataFormat = "h:mm";
  //                  break;
  //              case ExcelDateType.自定义11:
  //                  strDataFormat = "h:mm:ss";
  //                  break;
  //              case ExcelDateType.自定义12:
  //                  strDataFormat = "h时mm分";
  //                  break;
  //              case ExcelDateType.自定义13:
  //                  strDataFormat = "h时mm分ss秒";
  //                  break;
  //              case ExcelDateType.自定义14:
  //                  strDataFormat = "上午/下午h时mm分";
  //                  break;
  //              case ExcelDateType.自定义15:
  //                  strDataFormat = "上午/下午h时mm分ss秒";
  //                  break;
  //              case ExcelDateType.自定义16:
  //                  strDataFormat = "yyyy/m/d h:mm";
  //                  break;
  //              case ExcelDateType.自定义17:
  //                  strDataFormat = "mm:ss";
  //                  break;
  //              case ExcelDateType.自定义18:
  //                  strDataFormat = "mm:ss.0";
  //                  break;
  //          }
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置值：分数
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="workBook"></param>
  //      /// <param name="value"></param>
  //      /// <param name="facType"></param>
  //      /// <param name="cellStyle"></param>
  //      public static void SetCellValue(this ICell cell,  object value, ExcelFractionType facType )
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
            
  //          if (value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          string strDataFormat = string.Empty;
  //          switch (facType)
  //          {
  //              case ExcelFractionType.分数1:
  //                  strDataFormat = "# ?/?";
  //                  break;
  //              case ExcelFractionType.分数2:
  //                  strDataFormat = "# ??/??";
  //                  break;
  //              case ExcelFractionType.分数3:
  //                  strDataFormat = "# ???/???";
  //                  break;
  //              case ExcelFractionType.分数4:
  //                  strDataFormat = "# ?/2";
  //                  break;
  //              case ExcelFractionType.分数5:
  //                  strDataFormat = "# ?/4";
  //                  break;
  //              case ExcelFractionType.分数6:
  //                  strDataFormat = "# ?/8";
  //                  break;
  //              case ExcelFractionType.分数7:
  //                  strDataFormat = "# ?/16";
  //                  break;
  //              case ExcelFractionType.分数8:
  //                  strDataFormat = "# ?/10";
  //                  break;
  //              case ExcelFractionType.分数9:
  //                  strDataFormat = "# ??/100";
  //                  break;
  //          }
  //          DateTime dtDate = DateTime.Parse(value.ToString());
  //          cell.SetCellValue(dtDate);
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置值：特殊类型 数字转中文
  //      /// <para>先设置样式，再调用此赋值</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="workBook"></param>
  //      /// <param name="value"></param>
  //      /// <param name="specType"></param>
  //      /// <param name="cellStyle"></param>
  //      public static void SetCellValue(this ICell cell,  object value, ExcelSpecialType specType )
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
            
  //          if (value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          string strDataFormat = string.Empty;
  //          switch (specType)
  //          {
  //              case ExcelSpecialType.特殊1:
  //                  strDataFormat = "000000";
  //                  break;
  //              case ExcelSpecialType.特殊2:
  //                  strDataFormat = "[DBNum1][$-804]G/通用格式";
  //                  break;
  //              case ExcelSpecialType.特殊3:
  //                  strDataFormat = "[DBNum2][$-804]G/通用格式";
  //                  break;
  //          }
  //          double dbValue = 0;
  //          double.TryParse(value.ToString(), out dbValue);
  //          cell.SetCellValue(dbValue);
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat(strDataFormat);
  //          cell.CellStyle = cellStyle;

  //      }

  //      /// <summary>
  //      /// 设置值：自定义 DataFormat
  //      /// <para>先设置样式，再调用此赋值</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="value"></param>
  //      /// <param name="dataFormat"></param>
  //      private static void SetCellValue(this ICell cell, object value,string dataFormat)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
            
  //          if(value == null)
  //          {
  //              cell.SetCellValue(string.Empty);
  //              return;
  //          }
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat(dataFormat);
  //          cell.SetCellValue(value.ToString());
  //          cell.CellStyle = cellStyle;
  //      }

  //      #endregion

  //      #region 扩展单元格样式赋值

  //      /// <summary>
  //      /// 字体加粗
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      public static void SetBold(this ICell cell, FontBoldWeight bold = FontBoldWeight.Bold)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          IFont font = cellStyle.GetFont(workBook);
  //          font.Boldweight = (short)bold;
  //          cellStyle.SetFont(font);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 字体倾斜
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      public static void SetItalic(this ICell cell)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          IFont font = cellStyle.GetFont(workBook);
  //          font.IsItalic = true;
  //          cellStyle.SetFont(font);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置下划线
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="dbLine"></param>
  //      /// <param name="lineType"></param>
  //      public static void SetUnderline(this ICell cell, bool dbLine = false, FontUnderlineType lineType = FontUnderlineType.Single)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          IFont font = cellStyle.GetFont(workBook);
  //          font.Underline = lineType;
  //          cellStyle.SetFont(font);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置百分百：%
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgbs"></param>
  //      public static void SetPercent(this ICell cell)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();

  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat("0%");
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置千分位
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      public static void SetThousandsSeparator(this ICell cell)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          IDataFormat sdf = workBook.CreateDataFormat();
  //          cellStyle.DataFormat = sdf.GetFormat("_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ' - '??_ ;_ @_ ");
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置边框
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="boderType"></param>
  //      /// <param name="lineColor"></param>
  //      public static void SetBoderLine(this ICell cell, ExcelBorderType boderType = ExcelBorderType.BorderAll, ColorType lineColor = ColorType.black)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          string rgb = GetColor(lineColor).Item2;
  //          short indexed = workBook.GetCustomColor(rgb);
  //          switch (boderType)
  //          {
  //              case ExcelBorderType.BorderAll:
  //                  cellStyle.BorderBottom = BorderStyle.Thin;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  cellStyle.BorderTop = BorderStyle.Thin;
  //                  cellStyle.TopBorderColor = indexed;
  //                  cellStyle.BorderLeft = BorderStyle.Thin;
  //                  cellStyle.LeftBorderColor = indexed;
  //                  cellStyle.BorderRight = BorderStyle.Thin;
  //                  cellStyle.RightBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderAllBold:
  //                  cellStyle.BorderBottom = BorderStyle.Thick;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  cellStyle.BorderTop = BorderStyle.Thick;
  //                  cellStyle.TopBorderColor = indexed;
  //                  cellStyle.BorderLeft = BorderStyle.Thick;
  //                  cellStyle.LeftBorderColor = indexed;
  //                  cellStyle.BorderRight = BorderStyle.Thick;
  //                  cellStyle.RightBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderBottomBold:
  //                  cellStyle.BorderBottom = BorderStyle.Thick;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderBottom:
  //                  cellStyle.BorderBottom = BorderStyle.Thin;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderBottomDouble:
  //                  cellStyle.BorderBottom = BorderStyle.Double;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderLeft:
  //                  cellStyle.BorderLeft = BorderStyle.Thin;
  //                  cellStyle.LeftBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderRight:
  //                  cellStyle.BorderRight = BorderStyle.Thin;
  //                  cellStyle.RightBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderTop:
  //                  cellStyle.BorderTop = BorderStyle.Thin;
  //                  cellStyle.TopBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderTopAndBotoomBold:
  //                  cellStyle.BorderTop = BorderStyle.Thin;
  //                  cellStyle.TopBorderColor = indexed;
  //                  cellStyle.BorderBottom = BorderStyle.Thick;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderTopAndBottom:
  //                  cellStyle.BorderTop = BorderStyle.Thin;
  //                  cellStyle.TopBorderColor = indexed;
  //                  cellStyle.BorderBottom = BorderStyle.Thin;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //              case ExcelBorderType.BorderTopAndBottomDouble:
  //                  cellStyle.BorderTop = BorderStyle.Thin;
  //                  cellStyle.TopBorderColor = indexed;
  //                  cellStyle.BorderBottom = BorderStyle.Double;
  //                  cellStyle.BottomBorderColor = indexed;
  //                  break;
  //          }

  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 设置样式：Excel同样样式
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="csType"></param>
  //      public static void SetCellStyle(this ICell cell, ExcelCellStyleType csType)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          cellStyle = workBook.CreateCellStyle();
  //          IFont font = workBook.CreateFont();
  //          font.FontName = "宋体";
  //          font.FontHeightInPoints = 11;
  //          short indexed = defaultColorIndexed;
  //          switch (csType)
  //          {
  //              case ExcelCellStyleType.好:
  //                  font.Color = workBook.GetCustomColor("0,97,0");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("198,239,206");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.差:
  //                  font.Color = workBook.GetCustomColor("156,0,6");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("255,199,206");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.适中:
  //                  font.Color = workBook.GetCustomColor("156,101,0");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("255,235,156");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.计算:
  //                  font.Color = workBook.GetCustomColor("250,125,0");
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("242,242,242");
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderAll);
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色1:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("79,129,189");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色1_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("220,230,241");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色1_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("184,204,228");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色1_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("149,179,215");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色2:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("192,80,77");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色2_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("242,220,219");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色2_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("230,184,183");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色2_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("218,150,148");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色3:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("155,187,89");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色3_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("235,241,222");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色3_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("216,228,188");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色3_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("196,215,155");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色4:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("128,100,162");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色4_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("228,223,236");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色4_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("204,192,218");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色4_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("177,160,199");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色5:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("75,172,198");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色5_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("218,238,243");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色5_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("183,222,232");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色5_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("146,205,220");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色6:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("247,150,70");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色6_20:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("253,233,217");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色6_40:
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("252,213,180");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.强调文字颜色6_60:
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("250,191,143");
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.标题:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.FontHeightInPoints = 18;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.标题1:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.FontHeightInPoints = 15;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
  //                  break;
  //              case ExcelCellStyleType.标题2:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.FontHeightInPoints = 13;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
  //                  break;
  //              case ExcelCellStyleType.标题3:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
  //                  break;
  //              case ExcelCellStyleType.标题4:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.汇总:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderTopAndBottomDouble);
  //                  break;
  //              case ExcelCellStyleType.检查单元格:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.Color = workBook.GetCustomColor("255,255,255");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("165,165,165");
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderAll);
  //                  break;
  //              case ExcelCellStyleType.解释性文本:
  //                  font.Color = workBook.GetCustomColor("127,127,127");
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetItalic();
  //                  break;
  //              case ExcelCellStyleType.警告文本:
  //                  font.Color = workBook.GetCustomColor("255,0,0");
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  break;
  //              case ExcelCellStyleType.链接单元格:
  //                  font.Color = workBook.GetCustomColor("250,125,0");
  //                  cellStyle.SetFont(font);
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderBottomDouble);
  //                  break;
  //              case ExcelCellStyleType.输出:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.Color = workBook.GetCustomColor("63,63,63");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("242,242,242");
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderAll);
  //                  break;
  //              case ExcelCellStyleType.输入:
  //                  font.Boldweight = (short)FontBoldWeight.Bold;
  //                  font.Color = workBook.GetCustomColor("63,63,118");
  //                  cellStyle.SetFont(font);
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("255,204,153");
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderAll);
  //                  break;
  //              case ExcelCellStyleType.注释:
  //                  cellStyle.FillPattern = FillPattern.SolidForeground;
  //                  cellStyle.FillForegroundColor = workBook.GetCustomColor("255,255,204");
  //                  cell.CellStyle = cellStyle;
  //                  cell.SetBoderLine(ExcelBorderType.BorderAll);
  //                  break;
  //          }

  //      }

  //      /// <summary>
  //      /// 设置字体颜色
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="colorType"></param>
  //      public static void SetFontColor(this ICell cell, ColorType colorType)
  //      {
  //          if (cell == null)
  //              return;
  //          cell.DealParam();
  //          IFont font = workBook.GetFontAt(cellStyle.FontIndex);
  //          font.Color = workBook.GetCustomColor(colorType);
  //          cellStyle.SetFont(font);
  //          cell.CellStyle = cellStyle;
  //      }

  //      /// <summary>
  //      /// 填充颜色
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="color"></param>
  //      public static void FillColor(this ICell cell, ColorType color)
  //      {
  //          ICellStyle tempStyle = workBook.CreateCellStyle();
  //          tempStyle = cellStyle;

  //          tempStyle.FillPattern = FillPattern.SolidForeground;
  //          tempStyle.FillForegroundColor = workBook.GetCustomColor(color);
  //          cell.CellStyle = tempStyle;
  //      }

  //      /// <summary>
  //      /// 填充颜色
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      /// <param name="rgb"></param>
  //      public static void FillColor(this ICell cell, string rgb)
  //      {
  //          ICellStyle tempStyle = workBook.CreateCellStyle();
  //          tempStyle = cellStyle;
  //          tempStyle.FillPattern = FillPattern.SolidForeground;
  //          tempStyle.FillForegroundColor = workBook.GetCustomColor(rgb);
  //          cell.CellStyle = tempStyle;
  //      }

  //      #endregion

  //      #region 自定义颜色

  //      /// <summary>
  //      /// 设置自定义RGB颜色
  //      /// </summary>
  //      /// <param name="workBook">工作簿</param>
  //      /// <param name="rgbs">RGB颜色集合</param>
  //      public static void SetCustomColor(this IWorkbook workBook, IEnumerable<string> rgbs)
  //      {
  //          if (workBook is HSSFWorkbook)
  //          {
  //              // 设置颜色
  //              var tempWork = (HSSFWorkbook)workBook;
  //              tempWork.SetCustomColor(rgbs);
  //          }
  //      }

  //      /// <summary>
  //      /// 获取自定义颜色
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgb"></param>
  //      public static short GetCustomColor(this IWorkbook workBook, string rgb)
  //      {
  //          short indexed = defaultColorIndexed;
  //          if (workBook is HSSFWorkbook)
  //          {
  //              var tempWork = (HSSFWorkbook)workBook;
  //              indexed = tempWork.GetCustomColor(rgb);
  //          }
  //          else if (workBook is XSSFWorkbook)
  //          {
  //              var tempWork = (XSSFWorkbook)workBook;
  //              indexed = tempWork.GetCustomColor(rgb);
  //          }
  //          return indexed;
  //      }

  //      // 获取色值
  //      public static short GetCustomColor(this IWorkbook workBook, ColorType colorType)
  //      {
  //          short indexed = defaultColorIndexed;
  //          string rgb = GetColor(colorType).Item2;
  //          if(workBook is HSSFWorkbook)
  //          {
  //              var tempWork = (HSSFWorkbook)workBook;
  //              indexed = tempWork.GetCustomColor(rgb);
  //          }else if(workBook is XSSFWorkbook)
  //          {
  //              var tempWork = (XSSFWorkbook)workBook;
  //              indexed = tempWork.GetCustomColor(rgb);
  //          }
  //          return indexed;
  //      }

  //      /// <summary>
  //      /// 设置画板原生RGB颜色
  //      /// </summary>
  //      private static void SetOriginalRGB()
  //      {
  //          originalRGBs.Add(new Tuple<int, string>(8, "0,0,0"));
  //          originalRGBs.Add(new Tuple<int, string>(9, "255,255,255"));
  //          originalRGBs.Add(new Tuple<int, string>(10, "255,0,0"));
  //          originalRGBs.Add(new Tuple<int, string>(11, "0,255,0"));
  //          originalRGBs.Add(new Tuple<int, string>(12, "0,0,255"));
  //          originalRGBs.Add(new Tuple<int, string>(13, "255,0,0"));
  //          originalRGBs.Add(new Tuple<int, string>(14, "255,0,255"));
  //          originalRGBs.Add(new Tuple<int, string>(15, "0,255,255"));
  //          originalRGBs.Add(new Tuple<int, string>(16, "128,0,0"));
  //          originalRGBs.Add(new Tuple<int, string>(17,"0,128,0"));
  //          originalRGBs.Add(new Tuple<int, string>(18,"0,0,128"));
  //          originalRGBs.Add(new Tuple<int, string>(19,"128,128,0"));
  //          originalRGBs.Add(new Tuple<int, string>(20,"128,0,128"));
  //          originalRGBs.Add(new Tuple<int, string>(21,"0,128,128"));
  //          originalRGBs.Add(new Tuple<int, string>(22,"192,192,192"));
  //          originalRGBs.Add(new Tuple<int, string>(23,"128,128,128"));
  //          originalRGBs.Add(new Tuple<int, string>(24,"153,153,255"));
  //          originalRGBs.Add(new Tuple<int, string>(25,"153,51,102"));
  //          originalRGBs.Add(new Tuple<int, string>(26,"255,255,204"));
  //          originalRGBs.Add(new Tuple<int, string>(27,"204,255,255"));
  //          originalRGBs.Add(new Tuple<int, string>(28,"102,0,102"));
  //          originalRGBs.Add(new Tuple<int, string>(29,"255,128,128"));
  //          originalRGBs.Add(new Tuple<int, string>(30,"0,102,204"));
  //          originalRGBs.Add(new Tuple<int, string>(31,"204,204,255"));
  //          originalRGBs.Add(new Tuple<int, string>(32, "0,0,128"));
  //          originalRGBs.Add(new Tuple<int, string>(33,"255,0,255"));
  //          originalRGBs.Add(new Tuple<int, string>(34,"255,255,0"));
  //          originalRGBs.Add(new Tuple<int, string>(35,"0,255,255"));
  //          originalRGBs.Add(new Tuple<int, string>(36,"128,0,128"));
  //          originalRGBs.Add(new Tuple<int, string>(37,"128,0,0"));
  //          originalRGBs.Add(new Tuple<int, string>(38,"0,128,128"));
  //          originalRGBs.Add(new Tuple<int, string>(39,"0,0,255"));
  //          originalRGBs.Add(new Tuple<int, string>(40,"0,204,255"));
  //          originalRGBs.Add(new Tuple<int, string>(41,"204,255,255"));
  //          originalRGBs.Add(new Tuple<int, string>(42,"204,255,204"));
  //          originalRGBs.Add(new Tuple<int, string>(43,"255,255,153"));
  //          originalRGBs.Add(new Tuple<int, string>(44,"153,204,255"));
  //          originalRGBs.Add(new Tuple<int, string>(45,"255,153,204"));
  //          originalRGBs.Add(new Tuple<int, string>(46,"204,153,255"));
  //          originalRGBs.Add(new Tuple<int, string>(47,"255,204,153"));
  //          originalRGBs.Add(new Tuple<int, string>(48,"51,102,255"));
  //          originalRGBs.Add(new Tuple<int, string>(49,"51,204,204"));
  //          originalRGBs.Add(new Tuple<int, string>(50,"153,204,0"));
  //          originalRGBs.Add(new Tuple<int, string>(51,"255,204,0"));
  //          originalRGBs.Add(new Tuple<int, string>(52,"255,153,0"));
  //          originalRGBs.Add(new Tuple<int, string>(53,"255,102,0"));
  //          originalRGBs.Add(new Tuple<int, string>(54,"102,102,153"));
  //          originalRGBs.Add(new Tuple<int, string>(55,"150,150,150"));
  //          originalRGBs.Add(new Tuple<int, string>(56,"0,51,102"));
  //          originalRGBs.Add(new Tuple<int, string>(57,"51,153,102"));
  //          originalRGBs.Add(new Tuple<int, string>(58,"0,51,0"));
  //          originalRGBs.Add(new Tuple<int, string>(59,"51,51,0"));
  //          originalRGBs.Add(new Tuple<int, string>(60,"153,51,0"));
  //          originalRGBs.Add(new Tuple<int, string>(61,"153,51,102"));
  //          originalRGBs.Add(new Tuple<int, string>(62,"51,51,153"));
  //          originalRGBs.Add(new Tuple<int, string>(63, "51,51,51"));

  //      }

  //      #endregion

  //      #region 图片信息获取

  //      /// <summary>
  //      /// 获取重定义名称实体集合
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <returns></returns>
  //      public static IEnumerable<INameInfo> GetAllINameInfo(this IWorkbook workBook)
  //      {
  //          List<INameInfo> list = new List<INameInfo>();
  //          for (int index = 0; index < workBook.NumberOfNames; index++)
  //          {
  //              IName iname = workBook.GetNameAt(index);
  //              if (iname.IsDeleted || iname.IsFunctionName)
  //                  continue;
  //              INameInfo info = DealIName(iname);
  //              if (info == null)
  //                  continue;
  //              list.Add(info);
  //          }

  //          return list;
  //      }

  //      /// <summary>
  //      /// 获取所有图片信息
  //      /// 参考地址：https://www.cnblogs.com/hanzhaoxin/p/4442369.html
  //      /// </summary>
  //      /// <param name="sheet"></param>
  //      /// <returns></returns>
  //      public static IEnumerable<PictureInfo> GetAllPictureInfo(this ISheet sheet)
  //      {
  //          return sheet.GetAllPictureInfos(null, null, null, null);
  //      }
  //      #endregion

  //      #region 自定义颜色处理

  //      /// <summary>
  //      /// 低版本设置自定义颜色
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgbs"></param>
  //      private static void SetCustomColor(this HSSFWorkbook workBook, IEnumerable<string> rgbs)
  //      {
  //          SetOriginalRGB();
  //          // 获取调色板
  //          HSSFPalette pattern = workBook.GetCustomPalette();
  //          short indexed = 8;
            
  //          foreach (var rgb in rgbs)
  //          {
  //              if (indexed > 63)
  //                  indexed = 8;

  //              short tempIndex = pattern.SetCustomColor(rgb, indexed);
  //              if (tempIndex == -1)
  //                  continue;
  //              indexed = tempIndex;
  //              indexed++;
  //          }
  //      }

  //      /// <summary>
  //      /// 设置自定义颜色
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgb"></param>
  //      private static void SetCustomColor(this HSSFWorkbook workBook, string rgb)
  //      {
  //          SetOriginalRGB();

  //          HSSFPalette pattern = workBook.GetCustomPalette();
  //          pattern.SetCustomColor(rgb, -1);
  //      }

  //      /// <summary>
  //      /// 设置颜色
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="colorType"></param>
  //      private static void SetCustomColor(this HSSFWorkbook workBook, ColorType colorType)
  //      {
  //          string rgb = GetColor(colorType).Item2;
  //          workBook.SetCustomColor(rgb);
  //      }

  //      /// <summary>
  //      /// 设置颜色
  //      /// </summary>
  //      /// <param name="pattern"></param>
  //      /// <param name="rgb"></param>
  //      /// <param name="indexed"></param>
  //      /// <returns></returns>
  //      private static short SetCustomColor(this HSSFPalette pattern, string rgb, short indexed)
  //      {
  //          if (string.IsNullOrEmpty(rgb))
  //              return -1;
  //          string[] colors = rgb.Split(',');
  //          if (colors.Length != 3)
  //              return -1;
  //          byte red = 0;
  //          byte green = 0;
  //          byte blue = 0;
  //          // 处理RGB数据
  //          bool result = DealRGB(colors, ref red, ref green, ref blue);
  //          if (result == false)
  //              return -1;
  //          var temp = pattern.FindColor(red, green, blue);
  //          if (temp != null)
  //              return temp.Indexed;

  //          if (indexed == -1)
  //              indexed = 8;
  //          // 此位置下画板 原始rgb颜色
  //          string originalColor = originalRGBs.Where(t => t.Item1 == indexed).Select(t => t.Item2).FirstOrDefault();
  //          // 此位置下画板 rgb颜色
  //          string originalColor1 = string.Join(",", pattern.GetColor(indexed).RGB);
  //          // 如果两种颜色不一致，说明此位置已经设置了其他颜色，换个位置去设置
  //          if (originalColor != originalColor1)
  //          {
  //              indexed++;
  //              // 循环判断此位置颜色是否是原始颜色，如果是则设置，否则找其他位置
  //              // 如果此位置已经是最后位置了，则使用开始位置设置
  //              while (originalColor != originalColor1 || indexed < 64)
  //              {
  //                  originalColor = originalRGBs.Where(t => t.Item1 == indexed).Select(t => t.Item2).FirstOrDefault();
  //                  originalColor1 = string.Join(",", pattern.GetColor(indexed).RGB);
  //                  if (originalColor == originalColor1)
  //                      break;
  //                  indexed++;
  //              }
  //              if (indexed > 63)
  //                  indexed = 8;
  //          }

  //          pattern.SetColorAtIndex(indexed, red, green, blue);
  //          return indexed;
  //      }

  //      /// <summary>
  //      /// 
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgbs"></param>
  //      private static void SetCustomColor(this XSSFWorkbook workBook, IEnumerable<string> rgbs)
  //      {
  //          // 获取调色板
  //          XSSFColor color = null;
  //          short indexed = 8;
  //          foreach (var rgb in rgbs)
  //          {
  //              if (indexed > 63)
  //                  indexed = 8;
  //              if (string.IsNullOrEmpty(rgb))
  //                  continue;
  //              string[] colors = rgb.Split(',');
  //              if (colors.Length != 3)
  //                  continue;
  //              byte red = 0;
  //              byte green = 0;
  //              byte blue = 0;
  //              // 处理RGB数据
  //              bool result = DealRGB(colors, ref red, ref green, ref blue);
  //              if (result == false)
  //                  continue;
  //              byte[] bytes = { red, green, blue };
  //              color = new XSSFColor();
  //              color.SetRgb(bytes);
  //              color.Indexed = indexed;
  //              indexed++;
  //          }
  //      }

  //      /// <summary>
  //      /// 获取自定义颜色位置
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgb"></param>
  //      /// <returns></returns>
  //      private static short GetCustomColor(this HSSFWorkbook workBook, string rgb)
  //      {
  //          SetOriginalRGB();
  //          short indexed = defaultColorIndexed;
  //          if (string.IsNullOrEmpty(rgb))
  //              return indexed;
  //          string[] colors = rgb.Split(',');
  //          if (colors.Length != 3)
  //              return indexed;
  //          byte red = 0;
  //          byte green = 0;
  //          byte blue = 0;
  //          bool result = DealRGB(colors, ref red, ref green, ref blue);
  //          if (result == false)
  //              return indexed;
  //          HSSFPalette pattern = workBook.GetCustomPalette();
  //          NPOI.HSSF.Util.HSSFColor hssfColor = pattern.FindColor(red, green, blue);
  //          if (hssfColor == null)
  //              return pattern.SetCustomColor(rgb, -1);
  //          indexed = hssfColor.Indexed;
  //          return indexed;
  //      }

  //      /// <summary>
  //      /// 高版本获取自定义颜色
  //      /// </summary>
  //      /// <param name="workBook"></param>
  //      /// <param name="rgb"></param>
  //      /// <returns></returns>
  //      private static short GetCustomColor(this XSSFWorkbook workBook, string rgb)
  //      {
  //          short indexed = defaultColorIndexed;
  //          if (string.IsNullOrEmpty(rgb))
  //              return indexed;
  //          string[] colors = rgb.Split(',');
  //          if (colors.Length != 3)
  //              return indexed;
  //          byte red = 0;
  //          byte green = 0;
  //          byte blue = 0;
  //          bool result = DealRGB(colors, ref red, ref green, ref blue);
  //          if (result == false)
  //              return indexed;
  //          byte[] bytes = { red, green, blue };
  //          XSSFColor color = new XSSFColor();
  //          color.SetRgb(bytes);
  //          indexed = color.Indexed;

  //          return indexed;
  //      }

  //      /// <summary>
  //      /// 处理RGB
  //      /// </summary>
  //      /// <param name="colors"></param>
  //      /// <param name="red"></param>
  //      /// <param name="green"></param>
  //      /// <param name="blue"></param>
  //      private static bool DealRGB(string[] colors, ref byte red, ref byte green, ref byte blue)
  //      {
  //          bool result = true;
  //          red = 0;
  //          green = 0;
  //          blue = 0;
  //          if (byte.TryParse(colors[0], out red) &&
  //              byte.TryParse(colors[1], out green) &&
  //              byte.TryParse(colors[2], out blue))
  //          {
  //              // 如果超出255，则默认255；如果小于0，则默认0
  //              if (red > 255)
  //                  red = 255;
  //              if (red < 0)
  //                  red = 0;
  //              if (green > 255)
  //                  green = 255;
  //              if (green < 0)
  //                  green = 0;
  //              if (blue > 255)
  //                  blue = 255;
  //              if (blue < 0)
  //                  blue = 0;
  //          }
  //          else
  //              result = false;

  //          return result;
  //      }

  //      #endregion

  //      #region 重定义名称

  //      /// <summary>
  //      /// 处理位置
  //      /// 当前只处理单个单元格、或连续单元格
  //      /// 例如：Sheet2!$G$11 或 Sheet2!$D$5:$E$14 这两种情况
  //      /// </summary>
  //      /// <param name="iName"></param>
  //      /// <returns></returns>
  //      private static INameInfo DealIName(this IName iname)
  //      {
  //          if (string.IsNullOrEmpty(iname.RefersToFormula))
  //              return null;
  //          // 如果是跨区域则返回空； 例如：Sheet2!$D$5:$E$14,Sheet2!$G$11,Sheet2!$G$10
  //          string[] regions = iname.RefersToFormula.Split(',');
  //          if (regions.Length > 1)
  //              return null;
  //          INameInfo info = new INameInfo(iname);
  //          // 先替换不需要字符
  //          string region = iname.RefersToFormula.Replace(iname.SheetName, "").Replace("!", "").Replace("$", "");
  //          string[] postions = region.Split(':');

  //          int rowBegin = 0;
  //          int rowEnd = 0;
  //          int colBegin = 0;
  //          int colEnd = 0;
  //          if (postions.Length == 1) //单个单元格
  //          {
  //              // 处理位置
  //              DealPostion(postions[0], ref rowBegin, ref colBegin);
  //              // 位置
  //              info.FirstRow = info.LastRow = rowBegin;
  //              info.FirstCol = info.LastCol = colBegin;
  //          }
  //          else if (postions.Length == 2)//区域单元格
  //          {
  //              // 处理位置
  //              DealPostion(postions[0], ref rowBegin, ref colBegin);
  //              DealPostion(postions[1], ref rowEnd, ref colEnd);
  //              //位置赋值
  //              info.FirstRow = rowBegin;
  //              info.LastRow = rowEnd;
  //              info.FirstCol = colBegin;
  //              info.LastCol = colEnd;
  //          }
  //          if (info.FirstRow == info.LastRow)
  //              info.EqualRow = true;
  //          else
  //              info.EqualRow = false;

  //          if (info.FirstCol == info.LastCol)
  //              info.EqualCol = true;
  //          else
  //              info.EqualCol = false;
  //          return info;
  //      }

  //      /// <summary>
  //      /// 对区域进行处理，返回行位置和列位置
  //      /// </summary>
  //      /// <param name="postion"></param>
  //      /// <param name="rowIndex"></param>
  //      /// <param name="colIndex"></param>
  //      private static void DealPostion(string postion, ref int rowIndex, ref int colIndex)
  //      {
  //          if (postion.Length < 2)
  //              throw new Exception("invalid parameter");
  //          colIndex = ColumnNameToIndex(postion.Substring(0, 1));
  //          rowIndex = int.Parse(postion.Substring(1, 1)) - 1;
  //      }
  //      #endregion

  //      #region 处理图片获取

  //      /// <summary>
  //      /// 获取所有图片信息
  //      /// </summary>
  //      /// <param name="sheet"></param>
  //      /// <param name="minRow"></param>
  //      /// <param name="maxRow"></param>
  //      /// <param name="minCol"></param>
  //      /// <param name="maxCol"></param>
  //      /// <param name="onlyInternal"></param>
  //      /// <returns></returns>
  //      private static List<PictureInfo> GetAllPictureInfos(this ISheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal = true)
  //      {
  //          if (sheet is HSSFSheet)
  //          {
  //              return GetAllPictureInfos((HSSFSheet)sheet, firstRow, lastRow, firstCol, lastCol, onlyInternal);
  //          }
  //          else if (sheet is XSSFSheet)
  //          {
  //              return GetAllPictureInfos((XSSFSheet)sheet, firstRow, lastRow, firstCol, lastCol, onlyInternal);
  //          }
  //          else
  //          {
  //              throw new Exception("Unhandled type, Not added for this type: GetAllPicturesInfos() Extension method.");
  //          }
  //      }

  //      /// <summary>
  //      /// 获取所有图片信息：低版本
  //      /// </summary>
  //      /// <param name="sheet"></param>
  //      /// <param name="minRow"></param>
  //      /// <param name="maxRow"></param>
  //      /// <param name="minCol"></param>
  //      /// <param name="maxCol"></param>
  //      /// <param name="onlyInternal"></param>
  //      /// <returns></returns>
  //      private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal)
  //      {
  //          List<PictureInfo> picturesInfoList = new List<PictureInfo>();

  //          var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
  //          if (null != shapeContainer)
  //          {
  //              var shapeList = shapeContainer.Children;
  //              foreach (var shape in shapeList)
  //              {
  //                  if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
  //                  {
  //                      var picture = (HSSFPicture)shape;
  //                      var anchor = (HSSFClientAnchor)shape.Anchor;

  //                      if (IsInternalOrIntersect(firstRow, lastRow, firstCol, lastCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
  //                      {
  //                          picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.SuggestFileExtension(), picture.PictureData.Data));
  //                      }
  //                  }
  //              }
  //          }

  //          return picturesInfoList;
  //      }

  //      /// <summary>
  //      /// 获取所有图片信息：高版本
  //      /// </summary>
  //      /// <param name="sheet"></param>
  //      /// <param name="minRow"></param>
  //      /// <param name="maxRow"></param>
  //      /// <param name="minCol"></param>
  //      /// <param name="maxCol"></param>
  //      /// <param name="onlyInternal"></param>
  //      /// <returns></returns>
  //      private static List<PictureInfo> GetAllPictureInfos(XSSFSheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal)
  //      {
  //          List<PictureInfo> picturesInfoList = new List<PictureInfo>();

  //          var documentPartList = sheet.GetRelations();
  //          foreach (var documentPart in documentPartList)
  //          {
  //              if (documentPart is XSSFDrawing)
  //              {
  //                  var drawing = (XSSFDrawing)documentPart;
  //                  var shapeList = drawing.GetShapes();
  //                  foreach (var shape in shapeList)
  //                  {
  //                      if (shape is XSSFPicture)
  //                      {
  //                          var picture = (XSSFPicture)shape;
  //                          picture.PictureData.SuggestFileExtension();
  //                          var anchor = picture.GetPreferredSize();

  //                          if (IsInternalOrIntersect(firstRow, lastRow, firstCol, lastCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
  //                          {
  //                              picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.SuggestFileExtension(), picture.PictureData.Data));
  //                          }
  //                      }
  //                  }
  //              }
  //          }

  //          return picturesInfoList;
  //      }

  //      /// <summary>
  //      /// 判断
  //      /// </summary>
  //      /// <param name="rangeMinRow"></param>
  //      /// <param name="rangeMaxRow"></param>
  //      /// <param name="rangeMinCol"></param>
  //      /// <param name="rangeMaxCol"></param>
  //      /// <param name="pictureMinRow"></param>
  //      /// <param name="pictureMaxRow"></param>
  //      /// <param name="pictureMinCol"></param>
  //      /// <param name="pictureMaxCol"></param>
  //      /// <param name="onlyInternal"></param>
  //      /// <returns></returns>
  //      private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
  //          int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
  //      {
  //          int _rangeMinRow = rangeMinRow ?? pictureMinRow;
  //          int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
  //          int _rangeMinCol = rangeMinCol ?? pictureMinCol;
  //          int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

  //          if (onlyInternal)
  //          {
  //              return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
  //                      _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
  //          }
  //          else
  //          {
  //              return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
  //              (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
  //          }
  //      }

  //      #endregion

  //      #region 常用方法

  //      /// <summary>
  //      /// 单元格列名 转 列号
  //      /// </summary>
  //      /// <param name="columnName">列名</param>
  //      /// <returns></returns>
  //      private static int ColumnNameToIndex(string columnName)
  //      {
  //          if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

  //          int index = 0;
  //          char[] chars = columnName.ToUpper().ToCharArray();
  //          for (int i = 0; i < chars.Length; i++)
  //          {
  //              index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
  //          }
  //          return index - 1;
  //      }

  //      /// <summary>
  //      /// 单元格列号 转 列名
  //      /// </summary>
  //      /// <param name="index"></param>
  //      /// <returns></returns>
  //      private static string ColumnIndexToName(int index)
  //      {
  //          if (index < 0)
  //          {
  //              return null;
  //          }
  //          int num = 65;// A的Unicode码
  //          string colName = "";
  //          do
  //          {
  //              if (colName.Length > 0)
  //              {
  //                  index--;
  //              }
  //              int remainder = index % 26;
  //              colName = ((char)(remainder + num)) + colName;
  //              index = (int)((index - remainder) / 26);
  //          } while (index > 0);
  //          return colName;
  //      }

  //      /// <summary>
  //      /// 获取小数位
  //      /// </summary>
  //      /// <param name="dot">小数位</param>
  //      /// <returns></returns>
  //      private static string GetDot(int dot)
  //      {
  //          string strDot = string.Empty;
  //          if (dot == 0)
  //              return strDot;
  //          strDot = ".";
  //          for (int i = 0; i < dot; i++)
  //          {
  //              strDot += "0";
  //          }
  //          return strDot;
  //      }

  //      /// <summary>
  //      /// 处理参数
  //      /// <para>参数：IWorkBook, ISheet, ICellStyle</para>
  //      /// </summary>
  //      /// <param name="cell"></param>
  //      private static void DealParam(this ICell cell)
  //      {
  //          cellStyle = cell.CellStyle;
  //          sheet = cell.Sheet;
  //          workBook = sheet.Workbook;
  //          if (cellStyle == null)
  //              cellStyle = workBook.CreateCellStyle();
  //      }

  //      /// <summary>
  //      /// 返回设置颜色
  //      /// </summary>
  //      /// <param name="colorType"></param>
  //      /// <returns></returns>
  //      private static Tuple<string, string> GetColor(ColorType colorType)
  //      {
  //          Tuple<string, string> color = null;
  //          switch (colorType)
  //          {

  //              case ColorType.aliceblue:
  //                  color = new Tuple<string, string>("f0f8ff", "240,248,255");
  //                  break;
  //              case ColorType.antiquewhite:
  //                  color = new Tuple<string, string>("faebd7", "250,235,215");
  //                  break;
  //              case ColorType.aqua:
  //                  color = new Tuple<string, string>("00ffff", "0,255,255");
  //                  break;
  //              case ColorType.aquamarine:
  //                  color = new Tuple<string, string>("7fffd4", "127,255,212");
  //                  break;
  //              case ColorType.azure:
  //                  color = new Tuple<string, string>("f0ffff", "240,255,255");
  //                  break;
  //              case ColorType.beige:
  //                  color = new Tuple<string, string>("f5f5dc", "245,245,220");
  //                  break;
  //              case ColorType.bisque:
  //                  color = new Tuple<string, string>("ffe4c4", "255,228,196");
  //                  break;
  //              case ColorType.black:
  //                  color = new Tuple<string, string>("000000", "0,0,0");
  //                  break;
  //              case ColorType.blanchedalmond:
  //                  color = new Tuple<string, string>("ffebcd", "255,235,205");
  //                  break;
  //              case ColorType.blue:
  //                  color = new Tuple<string, string>("0000ff", "0,0,255");
  //                  break;
  //              case ColorType.blueviolet:
  //                  color = new Tuple<string, string>("8a2be2", "138,43,226");
  //                  break;
  //              case ColorType.brown:
  //                  color = new Tuple<string, string>("a52a2a", "165,42,42");
  //                  break;
  //              case ColorType.burlywood:
  //                  color = new Tuple<string, string>("deb887", "222,184,135");
  //                  break;
  //              case ColorType.cadetblue:
  //                  color = new Tuple<string, string>("5f9ea0", "95,158,160");
  //                  break;
  //              case ColorType.chartreuse:
  //                  color = new Tuple<string, string>("7fff00", "127,255,0");
  //                  break;
  //              case ColorType.chocolate:
  //                  color = new Tuple<string, string>("d2691e", "210,105,30");
  //                  break;
  //              case ColorType.coral:
  //                  color = new Tuple<string, string>("ff7f50", "255,127,80");
  //                  break;
  //              case ColorType.cornflowerblue:
  //                  color = new Tuple<string, string>("6495ed", "100,149,237");
  //                  break;
  //              case ColorType.cornsilk:
  //                  color = new Tuple<string, string>("fff8dc", "255,248,220");
  //                  break;
  //              case ColorType.crimson:
  //                  color = new Tuple<string, string>("dc143c", "220,20,60");
  //                  break;
  //              case ColorType.cyan:
  //                  color = new Tuple<string, string>("00ffff", "0,255,255");
  //                  break;
  //              case ColorType.darkblue:
  //                  color = new Tuple<string, string>("00008b", "0,0,139");
  //                  break;
  //              case ColorType.darkcyan:
  //                  color = new Tuple<string, string>("008b8b", "0,139,139");
  //                  break;
  //              case ColorType.darkgoldenrod:
  //                  color = new Tuple<string, string>("b8860b", "184,134,11");
  //                  break;
  //              case ColorType.darkgray:
  //                  color = new Tuple<string, string>("a9a9a9", "169,169,169");
  //                  break;
  //              case ColorType.darkgreen:
  //                  color = new Tuple<string, string>("006400", "0,100,0");
  //                  break;
  //              case ColorType.darkgrey:
  //                  color = new Tuple<string, string>("a9a9a9", "169,169,169");
  //                  break;
  //              case ColorType.darkkhaki:
  //                  color = new Tuple<string, string>("bdb76b", "189,183,107");
  //                  break;
  //              case ColorType.darkmagenta:
  //                  color = new Tuple<string, string>("8b008b", "139,0,139");
  //                  break;
  //              case ColorType.darkolivegreen:
  //                  color = new Tuple<string, string>("556b2f", "85,107,47");
  //                  break;
  //              case ColorType.darkorange:
  //                  color = new Tuple<string, string>("ff8c00", "255,140,0");
  //                  break;
  //              case ColorType.darkorchid:
  //                  color = new Tuple<string, string>("9932cc", "153,50,204");
  //                  break;
  //              case ColorType.darkred:
  //                  color = new Tuple<string, string>("8b0000", "139,0,0");
  //                  break;
  //              case ColorType.darksalmon:
  //                  color = new Tuple<string, string>("e9967a", "233,150,122");
  //                  break;
  //              case ColorType.darkseagreen:
  //                  color = new Tuple<string, string>("8fbc8f", "143,188,143");
  //                  break;
  //              case ColorType.darkslateblue:
  //                  color = new Tuple<string, string>("483d8b", "72,61,139");
  //                  break;
  //              case ColorType.darkslategray:
  //                  color = new Tuple<string, string>("2f4f4f", "47,79,79");
  //                  break;
  //              case ColorType.darkslategrey:
  //                  color = new Tuple<string, string>("2f4f4f", "47,79,79");
  //                  break;
  //              case ColorType.darkturquoise:
  //                  color = new Tuple<string, string>("00ced1", "0,206,209");
  //                  break;
  //              case ColorType.darkviolet:
  //                  color = new Tuple<string, string>("9400d3", "148,0,211");
  //                  break;
  //              case ColorType.deeppink:
  //                  color = new Tuple<string, string>("ff1493", "255,20,147");
  //                  break;
  //              case ColorType.deepskyblue:
  //                  color = new Tuple<string, string>("00bfff", "0,191,255");
  //                  break;
  //              case ColorType.dimgray:
  //                  color = new Tuple<string, string>("696969", "105,105,105");
  //                  break;
  //              case ColorType.dimgrey:
  //                  color = new Tuple<string, string>("696969", "105,105,105");
  //                  break;
  //              case ColorType.dodgerblue:
  //                  color = new Tuple<string, string>("1e90ff", "30,144,255");
  //                  break;
  //              case ColorType.firebrick:
  //                  color = new Tuple<string, string>("b22222", "178,34,34");
  //                  break;
  //              case ColorType.floralwhite:
  //                  color = new Tuple<string, string>("fffaf0", "255,250,240");
  //                  break;
  //              case ColorType.forestgreen:
  //                  color = new Tuple<string, string>("228b22", "34,139,34");
  //                  break;
  //              case ColorType.fuchsia:
  //                  color = new Tuple<string, string>("ff00ff", "255,0,255");
  //                  break;
  //              case ColorType.gainsboro:
  //                  color = new Tuple<string, string>("dcdcdc", "220,220,220");
  //                  break;
  //              case ColorType.ghostwhite:
  //                  color = new Tuple<string, string>("f8f8ff", "248,248,255");
  //                  break;
  //              case ColorType.gold:
  //                  color = new Tuple<string, string>("ffd700", "255,215,0");
  //                  break;
  //              case ColorType.goldenrod:
  //                  color = new Tuple<string, string>("daa520", "218,165,32");
  //                  break;
  //              case ColorType.gray:
  //                  color = new Tuple<string, string>("808080", "128,128,128");
  //                  break;
  //              case ColorType.green:
  //                  color = new Tuple<string, string>("008000", "0,128,0");
  //                  break;
  //              case ColorType.greenyellow:
  //                  color = new Tuple<string, string>("adff2f", "173,255,47");
  //                  break;
  //              case ColorType.grey:
  //                  color = new Tuple<string, string>("808080", "128,128,128");
  //                  break;
  //              case ColorType.honeydew:
  //                  color = new Tuple<string, string>("f0fff0", "240,255,240");
  //                  break;
  //              case ColorType.hotpink:
  //                  color = new Tuple<string, string>("ff69b4", "255,105,180");
  //                  break;
  //              case ColorType.indianred:
  //                  color = new Tuple<string, string>("cd5c5c", "205,92,92");
  //                  break;
  //              case ColorType.indigo:
  //                  color = new Tuple<string, string>("4b0082", "75,0,130");
  //                  break;
  //              case ColorType.ivory:
  //                  color = new Tuple<string, string>("fffff0", "255,255,240");
  //                  break;
  //              case ColorType.khaki:
  //                  color = new Tuple<string, string>("f0e68c", "240,230,140");
  //                  break;
  //              case ColorType.lavender:
  //                  color = new Tuple<string, string>("e6e6fa", "230,230,250");
  //                  break;
  //              case ColorType.lavenderblush:
  //                  color = new Tuple<string, string>("fff0f5", "255,240,245");
  //                  break;
  //              case ColorType.lawngreen:
  //                  color = new Tuple<string, string>("7cfc00", "124,252,0");
  //                  break;
  //              case ColorType.lemonchiffon:
  //                  color = new Tuple<string, string>("fffacd", "255,250,205");
  //                  break;
  //              case ColorType.lightblue:
  //                  color = new Tuple<string, string>("add8e6", "173,216,230");
  //                  break;
  //              case ColorType.lightcoral:
  //                  color = new Tuple<string, string>("f08080", "240,128,128");
  //                  break;
  //              case ColorType.lightcyan:
  //                  color = new Tuple<string, string>("e0ffff", "224,255,255");
  //                  break;
  //              case ColorType.lightgoldenrodyellow:
  //                  color = new Tuple<string, string>("fafad2", "250,250,210");
  //                  break;
  //              case ColorType.lightgray:
  //                  color = new Tuple<string, string>("d3d3d3", "211,211,211");
  //                  break;
  //              case ColorType.lightgreen:
  //                  color = new Tuple<string, string>("90ee90", "144,238,144");
  //                  break;
  //              case ColorType.lightgrey:
  //                  color = new Tuple<string, string>("d3d3d3", "211,211,211");
  //                  break;
  //              case ColorType.lightpink:
  //                  color = new Tuple<string, string>("ffb6c1", "255,182,193");
  //                  break;
  //              case ColorType.lightsalmon:
  //                  color = new Tuple<string, string>("ffa07a", "255,160,122");
  //                  break;
  //              case ColorType.lightseagreen:
  //                  color = new Tuple<string, string>("20b2aa", "32,178,170");
  //                  break;
  //              case ColorType.lightskyblue:
  //                  color = new Tuple<string, string>("87cefa", "135,206,250");
  //                  break;
  //              case ColorType.lightslategray:
  //                  color = new Tuple<string, string>("778899", "119,136,153");
  //                  break;
  //              case ColorType.lightslategrey:
  //                  color = new Tuple<string, string>("778899", "119,136,153");
  //                  break;
  //              case ColorType.lightsteelblue:
  //                  color = new Tuple<string, string>("b0c4de", "176,196,222");
  //                  break;
  //              case ColorType.lightyellow:
  //                  color = new Tuple<string, string>("ffffe0", "255,255,224");
  //                  break;
  //              case ColorType.lime:
  //                  color = new Tuple<string, string>("00ff00", "0,255,0");
  //                  break;
  //              case ColorType.limegreen:
  //                  color = new Tuple<string, string>("32cd32", "50,205,50");
  //                  break;
  //              case ColorType.linen:
  //                  color = new Tuple<string, string>("faf0e6", "250,240,230");
  //                  break;
  //              case ColorType.magenta:
  //                  color = new Tuple<string, string>("ff00ff", "255,0,255");
  //                  break;
  //              case ColorType.maroon:
  //                  color = new Tuple<string, string>("800000", "128,0,0");
  //                  break;
  //              case ColorType.mediumaquamarine:
  //                  color = new Tuple<string, string>("66cdaa", "102,205,170");
  //                  break;
  //              case ColorType.mediumblue:
  //                  color = new Tuple<string, string>("0000cd", "0,0,205");
  //                  break;
  //              case ColorType.mediumorchid:
  //                  color = new Tuple<string, string>("ba55d3", "186,85,211");
  //                  break;
  //              case ColorType.mediumpurple:
  //                  color = new Tuple<string, string>("9370db", "147,112,219");
  //                  break;
  //              case ColorType.mediumseagreen:
  //                  color = new Tuple<string, string>("3cb371", "60,179,113");
  //                  break;
  //              case ColorType.mediumslateblue:
  //                  color = new Tuple<string, string>("7b68ee", "123,104,238");
  //                  break;
  //              case ColorType.mediumspringgreen:
  //                  color = new Tuple<string, string>("00fa9a", "0,250,154");
  //                  break;
  //              case ColorType.mediumturquoise:
  //                  color = new Tuple<string, string>("48d1cc", "72,209,204");
  //                  break;
  //              case ColorType.mediumvioletred:
  //                  color = new Tuple<string, string>("c71585", "199,21,133");
  //                  break;
  //              case ColorType.midnightblue:
  //                  color = new Tuple<string, string>("191970", "25,25,112");
  //                  break;
  //              case ColorType.mintcream:
  //                  color = new Tuple<string, string>("f5fffa", "245,255,250");
  //                  break;
  //              case ColorType.mistyrose:
  //                  color = new Tuple<string, string>("ffe4e1", "255,228,225");
  //                  break;
  //              case ColorType.moccasin:
  //                  color = new Tuple<string, string>("ffe4b5", "255,228,181");
  //                  break;
  //              case ColorType.navajowhite:
  //                  color = new Tuple<string, string>("ffdead", "255,222,173");
  //                  break;
  //              case ColorType.navy:
  //                  color = new Tuple<string, string>("000080", "0,0,128");
  //                  break;
  //              case ColorType.oldlace:
  //                  color = new Tuple<string, string>("fdf5e6", "253,245,230");
  //                  break;
  //              case ColorType.olive:
  //                  color = new Tuple<string, string>("808000", "128,128,0");
  //                  break;
  //              case ColorType.olivedrab:
  //                  color = new Tuple<string, string>("6b8e23", "107,142,35");
  //                  break;
  //              case ColorType.orange:
  //                  color = new Tuple<string, string>("ffa500", "255,165,0");
  //                  break;
  //              case ColorType.orangered:
  //                  color = new Tuple<string, string>("ff4500", "255,69,0");
  //                  break;
  //              case ColorType.orchid:
  //                  color = new Tuple<string, string>("da70d6", "218,112,214");
  //                  break;
  //              case ColorType.palegoldenrod:
  //                  color = new Tuple<string, string>("eee8aa", "238,232,170");
  //                  break;
  //              case ColorType.palegreen:
  //                  color = new Tuple<string, string>("98fb98", "152,251,152");
  //                  break;
  //              case ColorType.paleturquoise:
  //                  color = new Tuple<string, string>("afeeee", "175,238,238");
  //                  break;
  //              case ColorType.palevioletred:
  //                  color = new Tuple<string, string>("db7093", "219,112,147");
  //                  break;
  //              case ColorType.papayawhip:
  //                  color = new Tuple<string, string>("ffefd5", "255,239,213");
  //                  break;
  //              case ColorType.peachpuff:
  //                  color = new Tuple<string, string>("ffdab9", "255,218,185");
  //                  break;
  //              case ColorType.peru:
  //                  color = new Tuple<string, string>("cd853f", "205,133,63");
  //                  break;
  //              case ColorType.pink:
  //                  color = new Tuple<string, string>("ffc0cb", "255,192,203");
  //                  break;
  //              case ColorType.plum:
  //                  color = new Tuple<string, string>("dda0dd", "221,160,221");
  //                  break;
  //              case ColorType.powderblue:
  //                  color = new Tuple<string, string>("b0e0e6", "176,224,230");
  //                  break;
  //              case ColorType.purple:
  //                  color = new Tuple<string, string>("800080", "128,0,128");
  //                  break;
  //              case ColorType.red:
  //                  color = new Tuple<string, string>("ff0000", "255,0,0");
  //                  break;
  //              case ColorType.rosybrown:
  //                  color = new Tuple<string, string>("bc8f8f", "188,143,143");
  //                  break;
  //              case ColorType.royalblue:
  //                  color = new Tuple<string, string>("4169e1", "65,105,225");
  //                  break;
  //              case ColorType.saddlebrown:
  //                  color = new Tuple<string, string>("8b4513", "139,69,19");
  //                  break;
  //              case ColorType.salmon:
  //                  color = new Tuple<string, string>("fa8072", "250,128,114");
  //                  break;
  //              case ColorType.sandybrown:
  //                  color = new Tuple<string, string>("f4a460", "244,164,96");
  //                  break;
  //              case ColorType.seagreen:
  //                  color = new Tuple<string, string>("2e8b57", "46,139,87");
  //                  break;
  //              case ColorType.seashell:
  //                  color = new Tuple<string, string>("fff5ee", "255,245,238");
  //                  break;
  //              case ColorType.sienna:
  //                  color = new Tuple<string, string>("a0522d", "160,82,45");
  //                  break;
  //              case ColorType.silver:
  //                  color = new Tuple<string, string>("c0c0c0", "192,192,192");
  //                  break;
  //              case ColorType.skyblue:
  //                  color = new Tuple<string, string>("87ceeb", "135,206,235");
  //                  break;
  //              case ColorType.slateblue:
  //                  color = new Tuple<string, string>("6a5acd", "106,90,205");
  //                  break;
  //              case ColorType.slategray:
  //                  color = new Tuple<string, string>("708090", "112,128,144");
  //                  break;
  //              case ColorType.slategrey:
  //                  color = new Tuple<string, string>("708090", "112,128,144");
  //                  break;
  //              case ColorType.snow:
  //                  color = new Tuple<string, string>("fffafa", "255,250,250");
  //                  break;
  //              case ColorType.springgreen:
  //                  color = new Tuple<string, string>("00ff7f", "0,255,127");
  //                  break;
  //              case ColorType.steelblue:
  //                  color = new Tuple<string, string>("4682b4", "70,130,180");
  //                  break;
  //              case ColorType.tan:
  //                  color = new Tuple<string, string>("d2b48c", "210,180,140");
  //                  break;
  //              case ColorType.teal:
  //                  color = new Tuple<string, string>("008080", "0,128,128");
  //                  break;
  //              case ColorType.thistle:
  //                  color = new Tuple<string, string>("d8bfd8", "216,191,216");
  //                  break;
  //              case ColorType.tomato:
  //                  color = new Tuple<string, string>("ff6347", "255,99,71");
  //                  break;
  //              case ColorType.turquoise:
  //                  color = new Tuple<string, string>("40e0d0", "64,224,208");
  //                  break;
  //              case ColorType.violet:
  //                  color = new Tuple<string, string>("ee82ee", "238,130,238");
  //                  break;
  //              case ColorType.wheat:
  //                  color = new Tuple<string, string>("f5deb3", "245,222,179");
  //                  break;
  //              case ColorType.white:
  //                  color = new Tuple<string, string>("ffffff", "255,255,255");
  //                  break;
  //              case ColorType.whitesmoke:
  //                  color = new Tuple<string, string>("f5f5f5", "245,245,245");
  //                  break;
  //              case ColorType.yellow:
  //                  color = new Tuple<string, string>("ffff00", "255,255,0");
  //                  break;
  //              case ColorType.yellowgreen:
  //                  color = new Tuple<string, string>("9acd32", "154,205,50");
  //                  break;
  //          }

  //          return color;
  //      }

  //      /// <summary>
		///// 判断是否包含
		///// </summary>
		///// <param name="str">原字符串</param>
		///// <param name="value">被包含字符串</param>
		///// <param name="ignoreCase">是否忽略大小写</param>
		///// <returns></returns>
		//private static bool Contains(string str, string value, bool ignoreCase = false)
  //      {
  //          bool flag = string.IsNullOrEmpty(str);
  //          bool result;
  //          if (flag)
  //          {
  //              result = false;
  //          }
  //          else
  //          {
  //              bool flag2 = string.IsNullOrEmpty(value);
  //              result = (flag2 || (ignoreCase ? str.ToLower().Contains(value.ToLower()) : str.Contains(value)));
  //          }
  //          return result;
  //      }

  //      /// <summary>
  //      /// 对Url进行编码
  //      /// </summary>
  //      /// <param name="url">url</param>
  //      /// <param name="encoding">字符编码</param>
  //      /// <param name="isUpper">编码字符是否转成大写,范例,"http://"转成"http%3A%2F%2F"</param>
  //      private static string UrlEncode(string url, Encoding encoding, bool isUpper = false)
  //      {
  //          string text = HttpUtility.UrlEncode(url, encoding);
  //          bool flag = !isUpper;
  //          string result;
  //          if (flag)
  //          {
  //              result = text;
  //          }
  //          else
  //          {
  //              result = GetUpperEncode(text);
  //          }
  //          return result;
  //      }

  //      /// <summary>
  //      /// 获取大写编码字符串
  //      /// </summary>
  //      private static string GetUpperEncode(string encode)
  //      {
  //          StringBuilder stringBuilder = new StringBuilder();
  //          int num = -2147483648;
  //          for (int i = 0; i < encode.Length; i++)
  //          {
  //              string text = encode[i].ToString();
  //              bool flag = text == "%";
  //              if (flag)
  //              {
  //                  num = i;
  //              }
  //              bool flag2 = i - num == 1 || i - num == 2;
  //              if (flag2)
  //              {
  //                  text = text.ToUpper();
  //              }
  //              stringBuilder.Append(text);
  //          }
  //          return stringBuilder.ToString();
  //      }
  //      #endregion

  //      #region Excel 单元格格式
  //      /**
  //       * G/通用格式
  //       * 0
  //       * 0.00
  //       * #,##0
  //       * #,##0.00
  //       * _ * #,##0_ ;_ * -#,##0_ ;_ * "-"_ ;_ @_ 
  //       * _ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ 
  //       * _ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* "-"_ ;_ @_ 
  //       * _ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ 
  //       * #,##0;-#,##0
  //       * #,##0;[Red]-#,##0
  //       * #,##0.00;-#,##0.00
  //       * #,##0.00;[Red]-#,##0.00
  //       * ¥#,##0;¥-#,##0
  //       * ¥#,##0;[Red]¥-#,##0
  //       * ¥#,##0.00;¥-#,##0.00
  //       * ¥#,##0.00;[Red]¥-#,##0.00
  //       * 0%
  //       * 0.00%
  //       * 0.00E+00
  //       * ##0.0E+0
  //       * # ?/?
  //       * # ??/??
  //       * 0.00_);[Red](0.00)
  //       * $#,##0_);($#,##0)
  //       * $#,##0_);[Red]($#,##0)
  //       * $#,##0.00_);($#,##0.00)
  //       * $#,##0.00_);[Red]($#,##0.00)
  //       * yyyy年m月
  //       * m月d日
  //       * yyyy/m/d
  //       * yyyy年m月d日
  //       * m/d/yy
  //       * d-mmm-yy
  //       * d-mmm
  //       * mmm-yy
  //       * h:mm AM/PM
  //       * h:mm:ss AM/PM
  //       * h:mm
  //       * h:mm:ss
  //       * h时mm分
  //       * h时mm分ss秒
  //       * 上午/下午h时mm分
  //       * 上午/下午h时mm分ss秒
  //       * yyyy/m/d h:mm
  //       * mm:ss
  //       * mm:ss.0
  //       * @
  //       * [h]:mm:ss
  //       * yyyy年m月d日
  //       * [$ZWL] #,##0.00;[Red][$ZWL] #,##0.00
  //       * _-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-
  //       * 
  //       * 
  //       * 
  //       * 
  //       **/
  //      #endregion
    }

}
