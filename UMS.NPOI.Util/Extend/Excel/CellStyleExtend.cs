using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    /// 单元格样式扩展
    /// </summary>
    public static class CellStyleExtend
    {
        private static IWorkbook workBook = null;
        private static ISheet sheet = null;
        private static ICellStyle cellStyle = null;
        private static readonly short defaultColorIndexed = 9;

        /// <summary>
        /// 字体加粗
        /// </summary>
        /// <param name="cell"></param>
        public static void SetBold(this ICell cell, FontBoldWeight bold = FontBoldWeight.Bold)
        {
            if (cell == null)
                return;
            cell.DealParam();
            IFont font = cellStyle.GetFont(workBook);
            font.Boldweight = (short)bold;
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 字体倾斜
        /// </summary>
        /// <param name="cell"></param>
        public static void SetItalic(this ICell cell)
        {
            if (cell == null)
                return;
            cell.DealParam();
            IFont font = cellStyle.GetFont(workBook);
            font.IsItalic = true;
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置下划线
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dbLine"></param>
        /// <param name="lineType"></param>
        public static void SetUnderline(this ICell cell, bool dbLine = false, FontUnderlineType lineType = FontUnderlineType.Single)
        {
            if (cell == null)
                return;
            cell.DealParam();
            IFont font = cellStyle.GetFont(workBook);
            font.Underline = lineType;
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置百分百：%
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgbs"></param>
        public static void SetPercent(this ICell cell)
        {
            if (cell == null)
                return;
            cell.DealParam();

            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat("0%");
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置千分位
        /// </summary>
        /// <param name="cell"></param>
        public static void SetThousandsSeparator(this ICell cell)
        {
            if (cell == null)
                return;
            cell.DealParam();
            IDataFormat sdf = workBook.CreateDataFormat();
            cellStyle.DataFormat = sdf.GetFormat("_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ' - '??_ ;_ @_ ");
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="boderType"></param>
        /// <param name="lineColor"></param>
        public static void SetBoderLine(this ICell cell, ExcelBorderType boderType = ExcelBorderType.BorderAll, ColorType lineColor = ColorType.black)
        {
            if (cell == null)
                return;
            cell.DealParam();
            string rgb = ExcelExtend.GetColor(lineColor).Item2;
            short indexed = workBook.GetCustomColor(rgb);
            switch (boderType)
            {
                case ExcelBorderType.BorderAll:
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BottomBorderColor = indexed;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.TopBorderColor = indexed;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.LeftBorderColor = indexed;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.RightBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderAllBold:
                    cellStyle.BorderBottom = BorderStyle.Thick;
                    cellStyle.BottomBorderColor = indexed;
                    cellStyle.BorderTop = BorderStyle.Thick;
                    cellStyle.TopBorderColor = indexed;
                    cellStyle.BorderLeft = BorderStyle.Thick;
                    cellStyle.LeftBorderColor = indexed;
                    cellStyle.BorderRight = BorderStyle.Thick;
                    cellStyle.RightBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderBottomBold:
                    cellStyle.BorderBottom = BorderStyle.Thick;
                    cellStyle.BottomBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderBottom:
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BottomBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderBottomDouble:
                    cellStyle.BorderBottom = BorderStyle.Double;
                    cellStyle.BottomBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderLeft:
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.LeftBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderRight:
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.RightBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderTop:
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.TopBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderTopAndBotoomBold:
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.TopBorderColor = indexed;
                    cellStyle.BorderBottom = BorderStyle.Thick;
                    cellStyle.BottomBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderTopAndBottom:
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.TopBorderColor = indexed;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BottomBorderColor = indexed;
                    break;
                case ExcelBorderType.BorderTopAndBottomDouble:
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.TopBorderColor = indexed;
                    cellStyle.BorderBottom = BorderStyle.Double;
                    cellStyle.BottomBorderColor = indexed;
                    break;
            }

            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置样式：Excel同样样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="csType"></param>
        public static void SetCellStyle(this ICell cell, ExcelCellStyleType csType)
        {
            if (cell == null)
                return;
            cell.DealParam();
            cellStyle = workBook.CreateCellStyle();
            IFont font = workBook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            short indexed = defaultColorIndexed;
            switch (csType)
            {
                case ExcelCellStyleType.好:
                    font.Color = workBook.GetCustomColor("0,97,0");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("198,239,206");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.差:
                    font.Color = workBook.GetCustomColor("156,0,6");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("255,199,206");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.适中:
                    font.Color = workBook.GetCustomColor("156,101,0");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("255,235,156");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.计算:
                    font.Color = workBook.GetCustomColor("250,125,0");
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("242,242,242");
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderAll);
                    break;
                case ExcelCellStyleType.强调文字颜色1:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("79,129,189");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色1_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("220,230,241");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色1_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("184,204,228");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色1_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("149,179,215");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色2:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("192,80,77");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色2_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("242,220,219");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色2_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("230,184,183");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色2_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("218,150,148");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色3:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("155,187,89");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色3_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("235,241,222");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色3_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("216,228,188");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色3_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("196,215,155");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色4:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("128,100,162");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色4_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("228,223,236");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色4_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("204,192,218");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色4_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("177,160,199");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色5:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("75,172,198");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色5_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("218,238,243");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色5_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("183,222,232");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色5_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("146,205,220");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色6:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("247,150,70");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色6_20:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("253,233,217");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色6_40:
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("252,213,180");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.强调文字颜色6_60:
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("250,191,143");
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.标题:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.FontHeightInPoints = 18;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.标题1:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.FontHeightInPoints = 15;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
                    break;
                case ExcelCellStyleType.标题2:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.FontHeightInPoints = 13;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
                    break;
                case ExcelCellStyleType.标题3:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderBottomBold);
                    break;
                case ExcelCellStyleType.标题4:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.汇总:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderTopAndBottomDouble);
                    break;
                case ExcelCellStyleType.检查单元格:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.Color = workBook.GetCustomColor("255,255,255");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("165,165,165");
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderAll);
                    break;
                case ExcelCellStyleType.解释性文本:
                    font.Color = workBook.GetCustomColor("127,127,127");
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetItalic();
                    break;
                case ExcelCellStyleType.警告文本:
                    font.Color = workBook.GetCustomColor("255,0,0");
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    break;
                case ExcelCellStyleType.链接单元格:
                    font.Color = workBook.GetCustomColor("250,125,0");
                    cellStyle.SetFont(font);
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderBottomDouble);
                    break;
                case ExcelCellStyleType.输出:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.Color = workBook.GetCustomColor("63,63,63");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("242,242,242");
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderAll);
                    break;
                case ExcelCellStyleType.输入:
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    font.Color = workBook.GetCustomColor("63,63,118");
                    cellStyle.SetFont(font);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("255,204,153");
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderAll);
                    break;
                case ExcelCellStyleType.注释:
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.FillForegroundColor = workBook.GetCustomColor("255,255,204");
                    cell.CellStyle = cellStyle;
                    cell.SetBoderLine(ExcelBorderType.BorderAll);
                    break;
            }

        }

        /// <summary>
        /// 设置字体颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="colorType"></param>
        public static void SetFontColor(this ICell cell, ColorType colorType)
        {
            if (cell == null)
                return;
            cell.DealParam();
            IFont font = workBook.GetFontAt(cellStyle.FontIndex);
            font.Color = workBook.GetCustomColor(colorType);
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 填充颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        public static void FillColor(this ICell cell, ColorType color)
        {
            ICellStyle tempStyle = workBook.CreateCellStyle();
            tempStyle = cellStyle;

            tempStyle.FillPattern = FillPattern.SolidForeground;
            tempStyle.FillForegroundColor = workBook.GetCustomColor(color);
            cell.CellStyle = tempStyle;
        }

        /// <summary>
        /// 填充颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rgb"></param>
        public static void FillColor(this ICell cell, string rgb)
        {
            ICellStyle tempStyle = workBook.CreateCellStyle();
            tempStyle = cellStyle;
            tempStyle.FillPattern = FillPattern.SolidForeground;
            tempStyle.FillForegroundColor = workBook.GetCustomColor(rgb);
            cell.CellStyle = tempStyle;
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
