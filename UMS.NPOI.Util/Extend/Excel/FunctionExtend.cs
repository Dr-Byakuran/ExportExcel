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
    /// 插入函数 扩展
    /// </summary>
    public static class FunctionExtend
    {
        #region 常用函数

        /// <summary>
        /// 求和
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Sum(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), beginRow + 1);
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), endRow + 1);
            cell.CellFormula = string.Concat("SUM(", begin, ":", end, ")");
        }

        /// <summary>
        /// 求和
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Sum(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), rowIndex + 1);
            cell.CellFormula = string.Concat("SUM(", begin, ")");
        }

        /// <summary>
        /// 统计
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Count(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), beginRow + 1);
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), endRow + 1);
            cell.CellFormula = string.Concat("COUNT(", begin, ":", end, ")");
        }

        /// <summary>
        /// 统计
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Count(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("COUNT(", begin, ")");
        }

        /// <summary>
        /// 平均值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Average(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), beginRow + 1);
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), endRow + 1);
            cell.CellFormula = string.Concat("COUNT(", begin, ":", end, ")");
        }

        /// <summary>
        /// 平均值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Average(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("AVERAGE(", begin, ")");
        }

        /// <summary>
        /// 最大值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Max(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), beginRow + 1);
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), endRow + 1);
            cell.CellFormula = string.Concat("MAX(", begin, ":", end, ")");
        }

        /// <summary>
        /// 最大值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Max(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("MAX(", begin, ")");
        }

        /// <summary>
        /// 最小值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Min(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), beginRow + 1);
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), endRow + 1);
            cell.CellFormula = string.Concat("MIN(", begin, ":", end, ")");
        }

        /// <summary>
        /// 最小值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Min(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("MIN(", begin, ")");
        }

        /// <summary>
        /// 标准差
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Stdev(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("STDEV(", begin, ":", end, ")");
        }

        #endregion

        #region 时间与日期
        
        /// <summary>
        /// 返回在 Microsoft Excel 日期时间代码中代表日期的数字
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        public static void Date(this ICell cell, int year, int month, int day)
        {
            cell.CellFormula = string.Concat("DATE(", year, ",", month, ",", day, ")");
        }

        /// <summary>
        /// 将日期值从字符串转换为序列数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText">例如：2019/01/02</param>
        public static void DateValue(this ICell cell, DateTime dateText)
        {
            cell.CellFormula = string.Concat("DATEVALUE(\"", dateText, "\")");
        }

        /// <summary>
        /// 获取日期在当前月所属第几天
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText">例如：2019/01/02</param>
        public static void Day(this ICell cell, DateTime dateText)
        {
            cell.CellFormula = string.Concat("DAY(\"", dateText, "\")");
        }

        /// <summary>
        /// 使用其他单元格值来计算日期在当前月所属第几天
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Day(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("DAY(", begin, ")");
        }

        /// <summary>
        /// 按一年360天，计算两个时间差时间
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateText">开始时间</param>
        /// <param name="endDateText">结束时间</param>
        public static void Days360(this ICell cell, DateTime beginDateText,DateTime endDateText)
        {
            cell.CellFormula = string.Concat("DAYS360(\"", beginDateText, "\",\"", endDateText, "\"");
        }

        /// <summary>
        /// 按一年360天，计算两个（单元格）时间差时间
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Days360(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("DAYS360(", begin, ",", end);
        }

        /// <summary>
        /// 按一年360天，计算两个时间差时间
        /// 一个是单元数据，一个是固定日期
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="flag">true：dateText作为结束日期，否则作为开始日期</param>
        public static void Days360(this ICell cell, DateTime dateText,int rowIndex, int colIndex, bool flag = true)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            if (flag)
                cell.CellFormula = string.Concat("DAYS360(", begin, ",\"", dateText, "\")");
            else
                cell.CellFormula = string.Concat("DAYS360(\"", dateText, "\",", begin, ")");
        }

        /// <summary>
        /// 在当前日期上加几个月
        /// <para>注意：得到的结果是一个日期序列</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText"></param>
        /// <param name="months"></param>
        public static void EDate(this ICell cell, DateTime dateText, int months)
        {
            cell.CellFormula = string.Concat("EDATE(\"", dateText, "\",", months, ")");
        }

        /// <summary>
        /// 在某个单元格数据上加几个月
        /// <para>注意：得到的结果是一个日期序列</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="months"></param>
        public static void EDate(this ICell cell,int rowIndex, int colIndex,int months)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("EDATE(", begin, ",", months, ")");
        }

        /// <summary>
        /// 表示指定月数之前或之后的月份的最后一天
        /// <para>注意：得到的结果是一个日期序列</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText"></param>
        /// <param name="months"></param>
        public static void EoMonth(this ICell cell, DateTime dateText, int months)
        {
            cell.CellFormula = string.Concat("EOMONTH(\"", dateText, "\",", months, ")");
        }

        /// <summary>
        /// 表示指定月数之前或之后的月份的最后一天
        /// <para>注意：得到的结果是一个日期序列</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="months"></param>
        public static void EoMonth(this ICell cell,int rowIndex, int colIndex, int months)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("EOMONTH(", begin, ",", months, ")");
        }

        /// <summary>
        /// 设置当前时间
        /// </summary>
        /// <param name="cell"></param>
        public static void Now(this ICell cell)
        {
            cell.CellFormula = "NOW()";
        }

        /// <summary>
        /// 当前日期
        /// </summary>
        /// <param name="cell"></param>
        public static void Today(this ICell cell)
        {
            cell.CellFormula = "TODAY()";
        }

        /// <summary>
        /// 年份
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Year(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("YEAR(", begin, ")");
        }

        /// <summary>
        /// 月份
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Month(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("MONTH(", begin, ")");
        }

        /// <summary>
        /// 小时
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Hour(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("HOUR(", begin, ")");
        }

        /// <summary>
        /// 分钟
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Minute(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("MINUTE(", begin, ")");
        }

        /// <summary>
        /// 秒
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Second(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("SECOND(", begin, ")");
        }
        
        /// <summary>
        /// 计算两个时间内工作日数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateText"></param>
        /// <param name="endDateText"></param>
        public static void NetWorkDays(this ICell cell, DateTime beginDateText, DateTime endDateText)
        {
            //NETWORKDAYS
            cell.CellFormula = string.Concat("NETWORKDAYS(\"", beginDateText, "\",\"", endDateText, "\")");
        }

        /// <summary>
        /// 计算两个时间内工作日数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void NetWorkDays(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("NETWORKDAYS(", begin, ",", end, ")");
        }

        /// <summary>
        /// 计算两个时间内工作日数 - 节假日
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <param name="hRow"></param>
        /// <param name="hCol"></param>
        public static void NetWorkDays(this ICell cell, int beginRow, int beginCol, int endRow, int endCol, int hRow, int hCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            string h = string.Concat(ExcelExtend.ColumnIndexToName(hCol), (hRow + 1));
            cell.CellFormula = string.Concat("NETWORKDAYS(", begin, ",", end, ",", h, ")");
        }

        /// <summary>
        /// 计算两个时间内工作日数 - 节假日集合
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <param name="beginHRow"></param>
        /// <param name="beginHCol"></param>
        /// <param name="endHRow"></param>
        /// <param name="endHCol"></param>
        public static void NetWorkDays(this ICell cell, int beginRow, int beginCol, int endRow, int endCol, int beginHRow, int beginHCol, int endHRow, int endHCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            string bh = string.Concat(ExcelExtend.ColumnIndexToName(beginHCol), (beginHRow + 1));
            string eh = string.Concat(ExcelExtend.ColumnIndexToName(endHCol), (endHRow + 1));
            cell.CellFormula = string.Concat("NETWORKDAYS(", begin, ",", end, ",", bh, ":", eh, ")");
        }

        /// <summary>
        /// 返回特定时间的序列数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="hour"></param>
        /// <param name="minute"></param>
        /// <param name="second"></param>
        public static void Time(this ICell cell, int hour, int minute, int second)
        {
            cell.CellFormula = string.Concat("TIME(", hour, ",", minute, ",", second, ")");
        }

        /// <summary>
        /// 返回当前日期代表一周中第几天
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText"></param>
        public static void WeekDay(this ICell cell, DateTime dateText)
        {
            cell.CellFormula = string.Concat("WEEKDAY(\"", dateText, "\")");
        }

        /// <summary>
        /// 返回当前日期代表一周中第几天
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void WeekDay(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("WEEKDAY(", begin, ")");
        }

        /// <summary>
        /// 返回当前日期所属当前年中第几周
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateText"></param>
        public static void WeekNum(this ICell cell, DateTime dateText)
        {
            cell.CellFormula = string.Concat("WEEKNUM(\"", dateText, "\")");
        }

        /// <summary>
        /// 返回当前日期所属当前年中第几周
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void WeekNum(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("WEEKNUM(", begin, ")");
        }

        /// <summary>
        /// 返回从开始时间 + N个工作日
        /// <para>注意：得到的是一个日期序列数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateText"></param>
        /// <param name="days"></param>
        public static void WorkDay(this ICell cell, DateTime beginDateText, int days)
        {
            cell.CellFormula = string.Concat("WEEKDAY(\"", beginDateText, "\",", days, ")");
        }

        /// <summary>
        /// 返回从开始时间 + N个工作日
        /// <para>注意：得到的是一个日期序列数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateRow"></param>
        /// <param name="beginDateCol"></param>
        /// <param name="days"></param>
        public static void WorkDay(this ICell cell, int beginDateRow, int beginDateCol, int days)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginDateCol), (beginDateRow + 1));
            cell.CellFormula = string.Concat("WEEKDAY(", begin, ",", days, ")");
        }

        /// <summary>
        /// 返回从开始时间 + N个工作日 - 节假日
        /// <para>注意：得到的是一个日期序列数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateRow"></param>
        /// <param name="beginDateCol"></param>
        /// <param name="days"></param>
        /// <param name="beginHRow"></param>
        /// <param name="beginHCol"></param>
        /// <param name="endHRow"></param>
        /// <param name="enndHCol"></param>
        public static void WorkDay(this ICell cell, DateTime dateText, int days, int beginHRow, int beginHCol, int endHRow, int enndHCol)
        {
            string bh = string.Concat(ExcelExtend.ColumnIndexToName(beginHCol), (beginHRow + 1));
            string eh = string.Concat(ExcelExtend.ColumnIndexToName(enndHCol), (endHRow + 1));
            cell.CellFormula = string.Concat("WORKDAY(\"", dateText, "\",", days, ",", bh, ":", eh, ")");
        }

        /// <summary>
        /// 返回从开始时间 + N个工作日 - 节假日
        /// <para>注意：得到的是一个日期序列数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginDateRow"></param>
        /// <param name="beginDateCol"></param>
        /// <param name="days"></param>
        /// <param name="beginHRow"></param>
        /// <param name="beginHCol"></param>
        /// <param name="endHRow"></param>
        /// <param name="enndHCol"></param>
        public static void WorkDay(this ICell cell, int beginDateRow, int beginDateCol, int days, int beginHRow, int beginHCol,int endHRow, int enndHCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginDateCol), (beginDateRow + 1));
            string bh = string.Concat(ExcelExtend.ColumnIndexToName(beginHCol), (beginHRow + 1));
            string eh = string.Concat(ExcelExtend.ColumnIndexToName(enndHCol), (endHRow + 1));
            cell.CellFormula = string.Concat("WORKDAY(", begin, ",", days, ",", bh, ":", eh, ")");
        }

        #endregion

        #region 数据与三角函数

        /// <summary>
        /// 绝对值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Abs(this ICell cell, int rowIndex, int colIndex)
        {
            string begin =  string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ABS(", begin, ")");
        }

        /// <summary>
        /// 设置绝对值
        /// </summary>
        /// <param name="cell"></param>
        public static void Abs(this ICell cell)
        {
            cell.CellFormula = string.Concat("ABS(\"", cell.ToString(), "\")");
        }

        /// <summary>
        /// 设置绝对值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Abs(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("ABS(", value, ")");
        }

        /// <summary>
        /// 返回一个弧度的反余弦
        /// </summary>
        /// <param name="cell"></param>
        public static void Acos(this ICell cell, double value)
        {
            if(value >= -1 && value <= 1)
            {
                cell.CellFormula = string.Concat("ACOS(", value, ")");
            }
        }

        /// <summary>
        /// 返回一个弧度的反余弦
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Acos(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ACOS(", begin, ")");
        }

        /// <summary>
        /// 返回反双曲余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Acosh(this ICell cell, double value)
        {
            if(value >= 1)
            {
                cell.CellFormula = string.Concat("ACOSH(", value, ")");
            }
        }

        /// <summary>
        /// 返回反双曲余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Acosh(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ACOSH(", begin, ")");
        }

        /// <summary>
        /// 返回一个有弧度的反正弦
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Asin(this ICell cell, double value)
        {
            if(value >= -1 && value <= 1)
            {
                cell.CellFormula = string.Concat("ASIN(", value, ")");
            }
        }

        /// <summary>
        /// 返回一个有弧度的反正弦
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Asin(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ASIN(", begin, ")");
        }

        /// <summary>
        /// 返回反双曲正弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Asinh(this ICell cell,double value)
        {
            if(value >= 1)
            {
                cell.CellFormula = string.Concat("ASINH(", value, ")");
            }
        }

        /// <summary>
        /// 返回反双曲正弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Asinh(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ASINH(", begin, ")");
        }

        /// <summary>
        /// 返回一个正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Atan(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("ATAN(", value, ")");
        }

        /// <summary>
        /// 返回一个正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Atan(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ATAN(", begin, ")");
        }

        /// <summary>
        /// 根据给定的X/Y轴坐标值，返回正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="x_num"></param>
        /// <param name="y_num"></param>
        public static void Atan2(this ICell cell, double x_num, double y_num)
        {
            cell.CellFormula = string.Concat("ATAN2(", x_num, ",", y_num, ")");
        }

        /// <summary>
        /// 根据给定的X/Y轴坐标值，返回正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="x_row"></param>
        /// <param name="x_col"></param>
        /// <param name="y_row"></param>
        /// <param name="y_col"></param>
        public static void Atan2(this ICell cell, int x_row, int x_col, int y_row, int y_col)
        {
            string x = string.Concat(ExcelExtend.ColumnIndexToName(x_col), (x_row + 1));
            string y = string.Concat(ExcelExtend.ColumnIndexToName(y_col), (y_row + 1));
            cell.CellFormula = string.Concat("ATAN2(", x, ",", y, ")");
        }

        /// <summary>
        /// 返回反双曲正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Atanh(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("ATANH(", value, ")");
        }

        /// <summary>
        /// 返回反双曲正切值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Atanh(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("ATANH(", begin, ")");
        }

        /// <summary>
        /// 返回给定角度的余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Cos(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("COS(", value, ")");
        }

        /// <summary>
        /// 返回给定角度的余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Cos(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("COS(", begin, ")");
        }

        /// <summary>
        /// 返回双曲余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Cosh(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("COSH(", value, ")");
        }

        /// <summary>
        /// 返回双曲余弦值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Cosh(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("COSH(", begin, ")");
        }

        /// <summary>
        /// 将弧度转换成角度
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Degrees(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("DEGREES(", value, ")");
        }

        /// <summary>
        /// 将弧度转换成角度
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Degress(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("DEGRESS(", begin, ")");
        }

        /// <summary>
        /// 将正数向上舍入到最近的偶数，负数向下舍入到最近的偶数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Even(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("EVEN(", value, ")");
        }

        /// <summary>
        /// 将正数向上舍入到最近的偶数，负数向下舍入到最近的偶数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Even(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("EVEN(", begin, ")");
        }

        /// <summary>
        /// 返回e的n次方
        /// <para>常数e = 2.71828182845904</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="n"></param>
        public static void Exp(this ICell cell, double n)
        {
            cell.CellFormula = string.Concat("EXP(", n, ")");
        }

        /// <summary>
        /// 返回e的n次方
        /// <para>常数e = 2.71828182845904</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Exp(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("EXP(", begin, ")");
        }

        /// <summary>
        /// 返回某数的阶乘
        /// <para>例如：1*2*..*n</para>
        /// <para>此数必须是非负数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Fact(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("FACT(", value, ")");
        }

        /// <summary>
        /// 返回某数的阶乘
        /// <para>例如：1*2*..*n</para>
        /// <para>此数必须是非负数</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Fact(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("FACT(", begin, ")");
        }

        /// <summary>
        /// 返回某数的双阶乘
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void FactDouble(this ICell cell ,double value)
        {
            cell.CellFormula = string.Concat("FACTDOUBLE(", value, ")");
        }

        /// <summary>
        /// 返回某数的双阶乘
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void FactDouble(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("FACTDOUBLE(", begin, ")");
        }

        /// <summary>
        /// 获取最大公约数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="values"></param>
        public static void Gcd(this ICell cell, IEnumerable<double> values)
        {
            cell.CellFormula = string.Concat("GCD(", string.Join(",", values), ")");
        }

        /// <summary>
        /// 获取最大公约数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Gcd(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("GCD(", begin, ":", end, ")");
        }

        /// <summary>
        /// 将数值向下取整为最近的整数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Int(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("INT(", value, ")");
        }

        /// <summary>
        /// 将数值向下取整为最近的整数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Int(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("INT(", begin, ")");
        }

        /// <summary>
        /// 返回最小公倍数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="values"></param>
        public static void Lcm(this ICell cell, IEnumerable<double> values)
        {
            cell.CellFormula = string.Concat("LCM(", string.Join(",", values), ")");
        }

        /// <summary>
        /// 返回最小公倍数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Lcm(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("LAC(", begin, ":", end, ")");
        }

        /// <summary>
        /// 返回给定数值的自然对数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Ln(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("LN(", value, ")");
        }

        /// <summary>
        /// 返回给定数值的自然对数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Ln(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("LN(", begin, ")");
        }

        /// <summary>
        /// 根据给定底数返回数字的对数
        /// <para>底数默认10</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="bs">底数</param>
        public static void Log(this ICell cell, double value, int bs = 10)
        {
            cell.CellFormula = string.Concat("LOG(", value, ",", bs, ")");
        }

        /// 根据给定底数返回数字的对数
        /// <para>底数默认10</para>
        public static void Log(this ICell cell, int rowIndex, int colIndex, int bs = 10)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("LOG(", begin, ",", bs, ")");
        }

        /// <summary>
        /// 返回一数组所代表的矩阵行列式的值
        /// <para>数组：行数和列数必须相等</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Moeterm(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            if(endRow - beginRow == endCol - beginCol)
            {
                string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
                string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
                cell.CellFormula = string.Concat("MDETERM(", begin, ":", end, ")");
            }
        }

        /// <summary>
        /// 返回一数组所代表的矩阵行列式的值
        /// <para>数组：行数和列数必须相等</para>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="arrs"></param>
        public static void Moeterm(this ICell cell, IEnumerable<IEnumerable<double>> arrs)
        {
            int lenRow = arrs.Count();
            bool result = true;
            List<string> list = new List<string>();
            arrs.ToList().ForEach(t =>
            {
                int len = t.Count();
                if(lenRow != len)
                {
                    result = false;
                    return;
                }
                list.Add(string.Join(",", t));
            });
            if (result)
            {
                cell.CellFormula = string.Concat("MOETERN({", string.Join(";", list), "})");
            }
        }

        /// <summary>
        /// 取余数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="num">被除数</param>
        /// <param name="divisor">除数</param>
        public static void Mod(this ICell cell, double num, double divisor)
        {
            cell.CellFormula = string.Concat("MOD(", num, ",", divisor, ")");
        }

        /// <summary>
        /// 取余数
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="numRow"></param>
        /// <param name="numCol"></param>
        /// <param name="divRow"></param>
        /// <param name="divCol"></param>
        public static void Mod(this ICell cell, int numRow, int numCol, int divRow, int divCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(numCol), (numRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(divCol), (divRow + 1));
            cell.CellFormula = string.Concat("MOD(", begin, ",", end, ")");
        }

        /// <summary>
        /// 返回某数的乘幂
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="num"></param>
        /// <param name="pow"></param>
        public static void Power(this ICell cell, double num, double pow)
        {
            cell.CellFormula = string.Concat("POWER(", num, ",", pow, ")");
        }

        /// <summary>
        /// 返回某数的乘幂
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="numRow"></param>
        /// <param name="numCol"></param>
        /// <param name="pow"></param>
        public static void Power(this ICell cell, int numRow, int numCol, double pow)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(numCol), (numRow + 1));
            cell.CellFormula = string.Concat("POWER(", begin, ",", pow, ")");
        }

        /// <summary>
        /// 计算所有参数乘积
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="values"></param>
        public static void Product(this ICell cell, IEnumerable<double> values)
        {
            cell.CellFormula = string.Concat("PRODUCT(", string.Join(",", values), ")");
        }

        /// <summary>
        /// 计算所有参数乘积
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Product(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("PRODUCT(", begin, ":", end, ")");
        }

        /// <summary>
        /// 返回0-1随机数
        /// </summary>
        /// <param name="cell"></param>
        public static void Rand(this ICell cell)
        {
            cell.CellFormula = string.Concat("RAND()");
        }

        /// <summary>
        /// 返回某数的平方根
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void Sqrt(this ICell cell, double value)
        {
            cell.CellFormula = string.Concat("SQRT(", value, ")");
        }

        /// <summary>
        /// 返回某数的平方根
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        public static void Sqrt(this ICell cell, int rowIndex, int colIndex)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(colIndex), (rowIndex + 1));
            cell.CellFormula = string.Concat("SQRT(", begin, ")");
        }

        /// <summary>
        /// 返回平方和
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="values"></param>
        public static void Sumsq(this ICell cell, IEnumerable<double> values)
        {
            cell.CellFormula = string.Concat("SUMSQ(", string.Join(",", values), ")");
        }

        /// <summary>
        /// 返回平方和
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="beginRow"></param>
        /// <param name="beginCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        public static void Sumsql(this ICell cell, int beginRow, int beginCol, int endRow, int endCol)
        {
            string begin = string.Concat(ExcelExtend.ColumnIndexToName(beginCol), (beginRow + 1));
            string end = string.Concat(ExcelExtend.ColumnIndexToName(endCol), (endRow + 1));
            cell.CellFormula = string.Concat("SUMSQ(", begin, ":", end, ")");
        }
        #endregion
    }
}
