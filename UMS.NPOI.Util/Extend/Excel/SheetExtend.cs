using NPOI.SS.UserModel;
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
    /// 工作表扩展
    /// </summary>
    public static class SheetExtend
    {
        /// <summary>
        /// 设置单独单元格 引用地址
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行位置：从0开始</param>
        /// <param name="colIndex">列位置：从0开始</param>
        public static string RefersToFormula(this ISheet sheet, int rowIndex, int colIndex)
        {
            string refer = string.Empty;
            refer = string.Concat(sheet.SheetName, "!", "$", ExcelExtend.ColumnIndexToName(colIndex), "$", (rowIndex + 1));
            return refer;
        }

        /// <summary>
        /// 获取区域单元格 引用地址
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="beginRow">开始行：从0开始</param>
        /// <param name="beginCol">开始列：从0开始</param>
        /// <param name="endRow">结束行：从0开始</param>
        /// <param name="endCol">结束列：从0开始</param>
        /// <returns></returns>
        public static string RefersToFormula(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            string refer = string.Empty;
            refer = string.Concat(sheet.SheetName, "!", "$", ExcelExtend.ColumnIndexToName(beginCol), "$", (beginRow + 1)) + ":" +
                    string.Concat("$", ExcelExtend.ColumnIndexToName(endCol), "$", (endRow + 1));
            return refer;
        }

        /// <summary>
        /// 单元格区域  文字加粗
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void SetBold(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for(var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for(var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetBold();
                }
            }
        }

        /// <summary>
        /// 单元格区域 字体倾斜
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void SetItalic(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetItalic();
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置下划线
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void SetUnderline(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetUnderline();
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置百分百：%
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgbs"></param>
        public static void SetPercent(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetPercent();
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置千分位
        /// </summary>
        /// <param name="cell"></param>
        public static void SetThousandsSeparator(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetThousandsSeparator();
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="boderType"></param>
        /// <param name="lineColor"></param>
        public static void SetBoderLine(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetBoderLine();
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置样式：Excel同样样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="csType"></param>
        public static void SetCellStyle(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol, ExcelCellStyleType type)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetCellStyle(type);
                }
            }
        }

        /// <summary>
        /// 单元格区域 设置字体颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="colorType"></param>
        public static void SetFontColor(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol, ColorType type)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.SetFontColor(type);
                }
            }
        }

        /// <summary>
        /// 单元格区域 填充颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        public static void FillColor(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol, ColorType type)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.FillColor(type);
                }
            }
        }

        /// <summary>
        /// 单元格区域 填充颜色
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rgb"></param>
        public static void FillColor(this ISheet sheet, int beginRow, int beginCol, int endRow, int endCol, string rgb)
        {
            IWorkbook workBook = sheet.Workbook;
            for (var rowIndex = beginRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (var colIndex = beginCol; colIndex <= endCol; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    cell.FillColor(rgb);
                }
            }
        }
    }
}
