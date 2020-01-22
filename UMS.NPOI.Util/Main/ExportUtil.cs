using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Data;
using System.IO;
using System.Web;
using System.Linq.Expressions;
using System.Reflection;
using UMS.Framework.NpoiUtil.Model;
using UMS.Framework.NpoiUtil;
using UMS.Framework.NpoiUtil.Util;
using NPOI.SS.Util;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil
{
    public class ExportUtil
    {
        private IWorkbook workBook = null;

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExportUtil()
        {

        }

        public string RunExport(ExportExcelType type, ExportRunEntity runEntity, params DataTable[] dtData)
        {
            string errorMessage = string.Empty;
            lock (this)
            {
                runEntity = ProcessCellStyle(runEntity);
                switch (type)
                {
                    case ExportExcelType.Simple:
                        errorMessage = RunExportSimple(runEntity, dtData);
                        break;
                    case ExportExcelType.Merge:
                        errorMessage = RunExportMerge(runEntity, dtData);
                        break;
                }
            }
            return errorMessage;
        }

        public string RunExport(ExportExcelType type, ExportRunEntity runEntity, DataTable dtMain, params DataTable[] dtSub)
        {
            string errorMessage = string.Empty;
            lock (this)
            {
                runEntity = ProcessCellStyle(runEntity);
                switch (type)
                {
                    case ExportExcelType.Reconciliation:
                        errorMessage = RunExportReconciliation(runEntity, dtMain, dtSub);
                        break;
                    case ExportExcelType.Bill:
                        errorMessage = RunExportBill(runEntity, dtMain, dtSub);
                        break;
                }
            }
            return errorMessage;
        }

        /// <summary>
        /// 简单列表数据导出
        /// 这里DataTable数据是已经排序过的，表头名称
        /// </summary>
        /// <returns></returns>
        private string RunExportSimple(ExportRunEntity helpEntity, params DataTable[] dtData)
        {
            /** 模板类型：标题行和内容行有边框， 其他网格线取消
             *  标题行     XX  XX  XX  XX
             *  内容行     xx  xx  xx  xx
             *  
             *  自定义内容 XX:xx     XX:xx
             * */
            string errorMessage = string.Empty;
            if (dtData == null)
                return "导出数据为空";
            foreach (var data in dtData)
            {
                if (data == null || data.Rows.Count == 0)
                {
                    errorMessage = "导出数据为空";
                    break;
                }
            }
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;
            if (helpEntity.ExportColumns == null || helpEntity.ExportColumns.Count() == 0)
                return "导出配置为空";
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;

            // 处理工作簿、文件名
            string fileName = BeforeDealSheet(helpEntity);
            // 处理每个工作表数据
            ExportSheetEntity sheetEntity = null;
            for (int i = 0; i < dtData.Length; i++)
            {
                sheetEntity = new ExportSheetEntity();
                if (i < helpEntity.SheetName.Count)
                    sheetEntity.SheetName = helpEntity.SheetName[i];
                if (i < helpEntity.SheetTitle.Count)
                    sheetEntity.SheetTitle = helpEntity.SheetTitle[i];
                errorMessage = BuildSinpleSheet(helpEntity, sheetEntity, dtData[i]);
                if (!string.IsNullOrEmpty(errorMessage))
                    break;
            }
            // 如果存在错误，则终止执行
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;

            // 写入Http
            HttpWrite(fileName);

            return errorMessage;
        }

        /// <summary>
        /// 导出标题合并类型
        /// </summary>
        /// <param name="dtData"></param>
        /// <param name="helpEntity"></param>
        /// <returns></returns>
        private string RunExportMerge(ExportRunEntity helpEntity, params DataTable[] dtData)
        {
            /** 模板类型：
             *  合并标题行                                   第5周                 第6周
             *  标题行 编号  名称  型号  采购周期  需求 总需求  即时库存  需求  总需求  即时库存
             *  内容行 xx     xx    xx     xx       xx    xx       xx      xx     xx       xx
             * */

            string errorMessage = string.Empty;
            if (dtData == null)
                return "导出数据为空";
            foreach (var dt in dtData)
            {
                if (dt == null || dt.Rows.Count == 0)
                {
                    errorMessage = "导出数据为空";
                    break;
                }
            }
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;
            if (helpEntity.ExportColumns == null || helpEntity.ExportColumns.Any() == false)
                return "导出配置为空";

            IEnumerable<string> mergeNames = helpEntity.ExportColumns.Where(t => string.IsNullOrEmpty(t.MergeName) == false).Select(t => t.MergeName).Distinct();
            if (mergeNames == null || mergeNames.Any() == false)
                return "合并标题为空";
            // 处理工作簿、文件名
            string fileName = BeforeDealSheet(helpEntity);

            // 处理每个工作表数据
            ExportSheetEntity sheetEntity = null;
            for (int i = 0; i < dtData.Length; i++)
            {
                sheetEntity = new ExportSheetEntity();
                if (i < helpEntity.SheetName.Count)
                    sheetEntity.SheetName = helpEntity.SheetName[i];
                if (i < helpEntity.SheetTitle.Count)
                    sheetEntity.SheetTitle = helpEntity.SheetTitle[i];
                sheetEntity.MergeNames = mergeNames;
                errorMessage = BuildMergeSheet(helpEntity, sheetEntity, dtData[i]);
                if (!string.IsNullOrEmpty(errorMessage))
                    break;
            }
            // 如果存在错误，则终止执行
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;

            // 写入Http
            HttpWrite(fileName);

            return errorMessage;
        }

        /// <summary>
        /// 对账单类型导出
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="dtMain"></param>
        /// <param name="dtSub"></param>
        /// <param name="helpEntity"></param>
        /// <returns></returns>
        private string RunExportReconciliation(ExportRunEntity helpEntity, DataTable dtMain, params DataTable[] dtSub)
        {
            /** 模板类型：标题行以上（含）需要冻结，使excel拉动时一直显示
             *                         xxx 对账单   （加上下划线）             
             *  供方：xxx      电话：xxx       传值：xxx     货币类别：xxx(需要加批准)
             *  客户：xxx      电话：xxx       传值：xxx     核对日期：xxx
             *                                 扣款金额：xxx 金额合计：xxx
             *  标题行     XXX     XXX     XXX     XXX     XXX     XXX     XXX     XXX
             *  内容行     xxx     xxx     xxx     xxx     xxx     xxx     xxx     xxx
             *  
             *  批准：如下
             *  提示:
                RMB:人民币,单价为含税价格
                USD:美元,单价为不含税价格
             * */
            #region 前期判判断和导出准备
            string errorMessage = string.Empty;
            if (dtMain == null || dtMain.Rows.Count == 0 || dtSub == null || (dtMain.Rows.Count != dtSub.Length))
                return "导出数据为空";
            if (helpEntity.ExportColumns == null || helpEntity.ExportColumns.Any() == false)
                return "配置数据为空";
            var masterColumn = helpEntity.ExportColumns.Where(t => t.PrimaryMark == true);
            var subColumn = helpEntity.ExportColumns.Where(t => t.PrimaryMark == false);
            if (masterColumn == null || masterColumn.Any() == false)
                return "主表配置数据为空";
            if (subColumn == null || subColumn.Any() == false)
                return "子表配置数据为空";
            int maxColIndex = masterColumn.Select(t => t.ColIndex).Max();
            int minColIndex = masterColumn.Select(t => t.ColIndex).Min();
            if (minColIndex == maxColIndex)
                return "主表配置列不能全在一个位置";

            #endregion
            var temp = masterColumn.Select(t => t.RowIndex).Distinct();
            foreach (var rowIndex in temp)
            {
                List<int> colIndexs = masterColumn.Where(t => t.RowIndex == rowIndex).OrderBy(t => t.ColIndex).Select(t => t.ColIndex).Distinct().ToList();
                bool result = EqualDiif(colIndexs, 1);
                if (result == false)
                {
                    errorMessage = "列间距必须为1";
                    break;
                }
            }
            if (!string.IsNullOrEmpty(errorMessage))
                return errorMessage;
            // 处理工作簿、文件名
            string fileName = BeforeDealSheet(helpEntity);
            ExportSheetEntity sheetEntity = null;
            for (int i = 0; i < dtMain.Rows.Count; i++)
            {
                var dtData = dtSub[i];
                sheetEntity = new ExportSheetEntity();
                if (i < helpEntity.SheetName.Count)
                    sheetEntity.SheetName = helpEntity.SheetName[i];
                if (i < helpEntity.SheetTitle.Count)
                    sheetEntity.SheetTitle = helpEntity.SheetTitle[i];
                BuildRecSheet(helpEntity, sheetEntity, dtMain.Rows[i], dtData);
            }

            // 写入Http
            HttpWrite(fileName);
            return errorMessage;
        }

        public string RunExportBill(ExportRunEntity helpEntity, DataTable dtMain, params DataTable[] dtSub)
        {
            string errorMessage = string.Empty;

            return errorMessage;
        }

        /// <summary>
        /// 处理文件名称，创建工作簿
        /// </summary>
        /// <param name="helpEntity"></param>
        /// <returns></returns>
        private string BeforeDealSheet(ExportRunEntity helpEntity)
        {
            // 文件名称
            string fileName = helpEntity.FileName;
            if (string.IsNullOrEmpty(fileName))
                fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            fileName = string.Concat(fileName, ".", helpEntity.Suffix.ToString().ToLower());

            // 默认创建低版本Excel工作簿
            switch (helpEntity.Suffix)
            {
                case ExportExcelSuffix.xls:
                    // 建立空白工作簿
                    workBook = new HSSFWorkbook();
                    break;
                case ExportExcelSuffix.xlsx:
                    // 建立空白工作簿
                    workBook = new XSSFWorkbook();
                    break;
                default:
                    // 建立空白工作簿
                    workBook = new HSSFWorkbook();
                    break;
            }
            return fileName;
        }

        /// <summary>
        /// 处理每个工作表数据
        /// </summary>
        /// <param name="dtData"></param>
        /// <param name="helpEntity"></param>
        /// <returns></returns>
        private string BuildSinpleSheet(ExportRunEntity helpEntity, ExportSheetEntity sheetEntity, DataTable dtData)
        {
            string errorMessage = string.Empty;
            // 在工作簿建立空白工作表
            ISheet sheet = null;
            if (!string.IsNullOrEmpty(sheetEntity.SheetName))
                sheet = workBook.CreateSheet(sheetEntity.SheetName);
            else
                sheet = workBook.CreateSheet();
            // 看是否有跳过
            int beginRow = 0 + helpEntity.SkipRowNum;
            int beginCol = 0 + helpEntity.SkipColNum;
            IRow rowHead = sheet.CreateRow(beginRow);
            // 循环添加表头
            int colIndex = beginCol;
            foreach (var item in helpEntity.ExportColumns)
            {
                if (item.Hidden) continue;
                ICell cell = rowHead.CreateCell(colIndex);
                cell.SetCellValue(item.ExcelName);
                cell.CellStyle = item.CellStyle;
                sheet.SetColumnWidth(colIndex, item.Width);
                rowHead.Height = helpEntity.THeight;
                IName iname = workBook.CreateName();
                iname.NameName = item.ColumnName;
                iname.RefersToFormula = string.Concat(sheet.SheetName, "!$", ExportExcelUtil.IndexToColName(colIndex), "$", beginRow + 1);
                colIndex++;
            }
            if (helpEntity.FreezeTitleRow)
                sheet.CreateFreezePane(0, 0 + helpEntity.SkipRowNum + 1, 0, helpEntity.SkipRowNum + 1);
            colIndex = beginCol;
            beginRow++;
            //循环赋值内容
            foreach (DataRow dr in dtData.Rows)
            {
                IRow rowContent = sheet.CreateRow(beginRow);
                ICell cell = null;
                colIndex = beginCol;
                foreach (var item in helpEntity.ExportColumns)
                {
                    if (item.Hidden) continue;
                    cell = rowContent.CreateCell(colIndex);
                    object curValue = dr[item.ColumnName];
                    cell.SetCellValue(curValue, item, item.CellStyle);
                    colIndex++;
                }
                rowContent.Height = (short)helpEntity.CHeight;
                beginRow++;
            }
            // 筛选
            if (helpEntity.AutoFilter)
            {
                CellRangeAddress c = new CellRangeAddress(0 + helpEntity.SkipRowNum, 0 + helpEntity.SkipRowNum, beginCol, colIndex);
                sheet.SetAutoFilter(c);
            }
            sheet.DisplayGridlines = helpEntity.ShowGridLine;
            ProcessSheet(sheet, helpEntity);
            return errorMessage;
        }

        private string BuildMergeSheet(ExportRunEntity helpEntity, ExportSheetEntity sheetEntity, DataTable dtData)
        {
            string errorMessage = string.Empty;

            // 设置开始行、列
            int beginRow = 0 + helpEntity.SkipRowNum;
            int beginCol = 0 + helpEntity.SkipColNum;
            ISheet sheet = workBook.CreateSheet();
            #region 标题处理
            IRow rowTitle1 = sheet.CreateRow(beginRow);
            rowTitle1.Height = helpEntity.THeight;
            int colIndex = beginCol;
            // 先赋值合并行标题
            foreach (var mergeName in sheetEntity.MergeNames)
            {
                var temp = helpEntity.ExportColumns.Where(t => t.MergeName == mergeName);
                var len = temp.Count();
                // 获取合并标题开始列位置
                int index = GetIndex(helpEntity.ExportColumns, mergeName);
                colIndex = index;
                ICell cellMerge = rowTitle1.CreateCell(colIndex);
                cellMerge.SetCellValue(mergeName);
                //cellMerge.CellStyle = cellStyleHead;
                // 设置合并
                sheet.AddMergedRegion(new CellRangeAddress(beginRow, beginRow, colIndex, colIndex + len - 1));
                //for (int tempIndex = colIndex; tempIndex < colIndex + len; tempIndex++)
                //    HSSFCellUtil.GetCell(rowTitle1, tempIndex).CellStyle = cellStyleHead;
            }
            beginRow++;
            IRow rowTitle2 = sheet.CreateRow(beginRow);
            rowTitle2.Height = helpEntity.THeight;
            colIndex = beginCol;
            foreach (var item in helpEntity.ExportColumns)
            {
                ICell cellTitle = rowTitle2.CreateCell(colIndex);
                cellTitle.SetCellValue(item.ExcelName);
                cellTitle.CellStyle = item.CellStyle;
                // 如果此标题在上述没有合并，则向上合并一列，保持美观
                if (string.IsNullOrEmpty(item.MergeName))
                {
                    sheet.AddMergedRegion(new CellRangeAddress(beginRow - 1, beginRow, colIndex, colIndex));
                    ICell cellTemp = HSSFCellUtil.GetCell(rowTitle1, colIndex);
                    //cellTemp.CellStyle = cellStyleHead;
                    cellTemp.SetCellValue(item.ExcelName);
                }
                sheet.SetColumnWidth(colIndex, item.Width);
                colIndex++;
            }
            // 冻结标题行
            if (helpEntity.FreezeTitleRow)
                sheet.CreateFreezePane(beginCol + helpEntity.MergeColNum, 2, beginCol + helpEntity.MergeColNum, 2);
            #endregion

            beginRow++;
            // 循环赋值列表内容
            foreach (DataRow dr in dtData.Rows)
            {
                IRow rowContent = sheet.CreateRow(beginRow);
                rowContent.Height = helpEntity.CHeight;
                colIndex = beginCol;
                ICell cell = null;
                // 赋值内容
                foreach (var item in helpEntity.ExportColumns)
                {
                    if (item.Hidden) continue;
                    cell = rowContent.CreateCell(colIndex);
                    object curValue = dr[item.ColumnName];
                    cell.SetCellValue(curValue, item, item.CellStyle);
                    colIndex++;
                }
                rowContent.Height = helpEntity.CHeight;
                beginRow++;
            }
            // 筛选
            if (helpEntity.AutoFilter)
            {
                CellRangeAddress c = new CellRangeAddress(0 + helpEntity.SkipRowNum + 1, 0 + helpEntity.SkipRowNum + 1, beginCol, colIndex);
                sheet.SetAutoFilter(c);
            }
            sheet.DisplayGridlines = helpEntity.ShowGridLine;
            ProcessSheet(sheet, helpEntity);
            return errorMessage;
        }

        private string BuildRecSheet(ExportRunEntity helpEntity, ExportSheetEntity sheetEntity, DataRow drMain, DataTable dtSub)
        {
            string errorMessage = string.Empty;
            ISheet sheet = null;
            if (!string.IsNullOrEmpty(sheetEntity.SheetName))
                sheet = workBook.CreateSheet(sheetEntity.SheetName);
            else
                sheet = workBook.CreateSheet();
            // 设置开始行和列
            int beginRow = 0 + helpEntity.SkipRowNum;
            int beginCol = 0 + helpEntity.SkipColNum;

            #region 对表头数据进行赋值
            // 标题行
            IRow rowTitle = sheet.CreateRow(beginRow);
            ICell cellTitle = rowTitle.CreateCell(beginCol);
            cellTitle.SetCellValue(sheetEntity.SheetTitle);
            // 设置下划线
            IFont fontLine = workBook.CreateFont();
            fontLine.Underline = FontUnderlineType.Single;
            if (helpEntity.TitleBoldMark == true)
                fontLine.FontHeight = (double)FontBoldWeight.Bold;
            cellTitle.CellStyle.SetFont(fontLine);
            cellTitle.CellStyle.Alignment = HorizontalAlignment.Center;
            cellTitle.CellStyle.VerticalAlignment = VerticalAlignment.Center;
            rowTitle.Height = helpEntity.THeight;
            // 标题行合并
            sheet.AddMergedRegion(new CellRangeAddress(beginRow, beginRow, beginCol, beginCol + helpEntity.ExportColumns.Count() - 1));
            beginRow++;
            // 表头字段集合
            var masterColumn = helpEntity.ExportColumns.Where(t => t.PrimaryMark == true);
            var subColumn = helpEntity.ExportColumns.Where(t => t.PrimaryMark == false);
            int maxColIndex = masterColumn.Select(t => t.ColIndex).Max();
            int minColIndex = masterColumn.Select(t => t.ColIndex).Min();
            var listTitleContent = masterColumn;
            var temp = masterColumn.Select(t => t.RowIndex).Distinct();
            int colIndex = beginCol;
            // 循环赋值表头数据
            foreach (var rowIndex in temp)
            {
                IRow rowContent = sheet.CreateRow(beginRow);
                var curRow = listTitleContent.Where(t => t.RowIndex == rowIndex).OrderBy(t => t.ColIndex);
                colIndex = beginCol;
                var curColMinIndex = curRow.Select(t => t.ColIndex).Min();
                var diff = curColMinIndex - minColIndex;
                // 此处判断主要是为了有跳过列的
                if (diff > 0)
                {
                    // 标题
                    colIndex += diff;
                    // 内容
                    colIndex += diff;
                }
                foreach (var item in curRow)
                {
                    colIndex = colIndex - item.diffNum;
                    // 先赋值标题，再赋值
                    ICell curTitle = rowContent.CreateCell(colIndex);
                    curTitle.SetCellValue(item.ExcelName);
                    curTitle.CellStyle = item.CellStyle;
                    IName iname = workBook.CreateName();
                    iname.NameName = item.ColumnName;
                    iname.RefersToFormula = string.Concat(sheet.SheetName, "!$", ExportExcelUtil.IndexToColName(colIndex), "$", beginRow + 1);
                    colIndex++;
                    ICell cell = rowContent.CreateCell(colIndex);
                    string curColumn = item.ColumnName;
                    object curValue = drMain[curColumn];
                    if (item.TitleColSpan > 1)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(beginRow, beginRow, colIndex, colIndex + item.TitleColSpan - 1));
                        colIndex++;
                    }
                    cell.SetCellValue(curValue, item, item.CellStyle);
                    colIndex++;
                }
                beginRow++;
            }
            #endregion
            // 获取和循环设置表体表体行
            var listTitle = helpEntity.ExportColumns;
            IRow rowSubTitle = sheet.CreateRow(beginRow);
            colIndex = beginCol;
            foreach (var item in listTitle)
            {
                ICell cellSubTitle = rowSubTitle.CreateCell(colIndex);
                cellSubTitle.SetCellValue(item.ExcelName);
                sheet.SetColumnWidth(colIndex, item.Width);
                colIndex++;
            }
            beginRow++;
            // 冻结上述行和列
            if (helpEntity.FreezeTitleRow)
                sheet.CreateFreezePane(beginCol, beginRow - 1, beginCol, beginRow - 1);
            // 循环赋值列表数据
            foreach (DataRow dr in dtSub.Rows)
            {
                IRow rowSubContent = sheet.CreateRow(beginRow);
                colIndex = beginCol;
                foreach (var item in listTitle)
                {
                    ICell cellSubContent = rowSubContent.CreateCell(colIndex);
                    object curValue = dr[item.ColumnName];
                    // 设置表体内容
                    cellSubContent.SetCellValue(curValue, item, item.CellStyle);
                    colIndex++;
                }
                beginRow++;
            }
            // 筛选
            if (helpEntity.AutoFilter)
            {
                CellRangeAddress c = new CellRangeAddress(0 + helpEntity.SkipRowNum + 1, 0 + helpEntity.SkipRowNum + 1, beginCol, colIndex);
                sheet.SetAutoFilter(c);
            }
            sheet.DisplayGridlines = helpEntity.ShowGridLine;
            ProcessSheet(sheet, helpEntity);
            return errorMessage;
        }

        private string BuildBillSheet(ExportRunEntity helpEntity, ExportSheetEntity sheetEntity, DataTable dtMain, DataTable[] dtSub)
        {
            string errorMessage = string.Empty;
            // 在工作簿建立空白工作表
            ISheet sheet = null;

            if (!string.IsNullOrEmpty(sheetEntity.SheetName))
                sheet = workBook.CreateSheet(sheetEntity.SheetName);
            else
                sheet = workBook.CreateSheet();
            // 看是否有跳过
            int beginRow = 0 + helpEntity.SkipRowNum;
            int beginCol = 0 + helpEntity.SkipColNum;
            IRow rowHead = sheet.CreateRow(beginRow);
            // 循环添加表头
            int colIndex = beginCol;
            var tempMast = helpEntity.ExportColumns.Where(t => t.PrimaryMark == true);
            var tempSub = helpEntity.ExportColumns.Where(t => t.PrimaryMark == false);
            var tempCol = tempMast.Union(tempSub);
            foreach (var item in tempCol)
            {
                if (item.Hidden) continue;
                ICell cell = rowHead.CreateCell(colIndex);
                cell.SetCellValue(item.ExcelName);
                cell.CellStyle = item.CellStyle;
                sheet.SetColumnWidth(colIndex, item.Width);
                rowHead.Height = helpEntity.THeight;
                IName iname = workBook.CreateName();
                iname.NameName = item.ColumnName;
                iname.RefersToFormula = string.Concat(sheet.SheetName, "!$", ExportExcelUtil.IndexToColName(colIndex), "$", beginRow + 1);
                colIndex++;
            }
            if (helpEntity.FreezeTitleRow)
                sheet.CreateFreezePane(0, 0 + helpEntity.SkipRowNum + 1, 0, helpEntity.SkipRowNum + 1);
            colIndex = beginCol;
            beginRow++;
            //循环赋值内容
            for (int i = 0; i < dtMain.Rows.Count; i++)
            {
                DataRow drMain = dtMain.Rows[i];
                DataTable dtData = dtSub[i];
                for (int j = 0; j < dtData.Rows.Count; j++)
                {
                    DataRow drSub = dtData.Rows[j];
                    IRow row = sheet.CreateRow(beginRow);
                    ICell cell = null;
                    colIndex = beginCol;
                    foreach (var col in tempCol)
                    {
                        if (col.Hidden) continue;
                        cell = row.CreateCell(colIndex);
                        object value = null;
                        if (col.PrimaryMark == true)
                        {
                            value = drMain[col.ColumnName];
                            if (helpEntity.OneMain == true && j > 0)
                            {
                                colIndex++;
                                continue;
                            }
                        }
                        value = drSub[col.ColumnName];
                        cell.SetCellValue(value, col, col.CellStyle);
                        colIndex++;
                    }
                    beginRow++;
                }
            }
            // 筛选
            if (helpEntity.AutoFilter)
            {
                CellRangeAddress c = new CellRangeAddress(0 + helpEntity.SkipRowNum, 0 + helpEntity.SkipRowNum, beginCol, colIndex);
                sheet.SetAutoFilter(c);
            }
            sheet.DisplayGridlines = helpEntity.ShowGridLine;
            ProcessSheet(sheet, helpEntity);
            return errorMessage;
        }

        private ExportRunEntity ProcessCellStyle(ExportRunEntity helpEntity)
        {
            IEnumerable<ColorEntity> listColors = ProcessColor(helpEntity);
            List<ICellStyle> listStyle = new List<ICellStyle>();
            ICellStyle cellStyle = null;
            List<CellStyleEntity> listCellStyle = new List<CellStyleEntity>();
            CellStyleEntity styleEntity = null;
            foreach (var item in helpEntity.ExportStyles)
            {
                cellStyle = workBook.CreateCellStyle();
                cellStyle.Alignment = item.Alignment;
                cellStyle.BorderBottom = item.BorderBottom;
                cellStyle.BorderDiagonal = item.BorderDiagonal;
                var temp = listColors.Where(t => t.RGB == item.BorderDiagonalColor);
                if (temp != null && temp.Any())
                    cellStyle.BorderDiagonalColor = temp.FirstOrDefault().Index;
                cellStyle.BorderDiagonalLineStyle = item.BorderDiagonalLineStyle;
                cellStyle.BorderLeft = item.BorderLeft;
                cellStyle.BorderRight = item.BorderRight;
                cellStyle.BorderTop = item.BorderTop;
                temp = listColors.Where(t => t.RGB == item.BottomBorderColor);
                if (temp != null && temp.Any())
                    cellStyle.BottomBorderColor = temp.FirstOrDefault().Index;

                if (!string.IsNullOrEmpty(item.DataFormat))
                {
                    IDataFormat df = workBook.CreateDataFormat();
                    cellStyle.DataFormat = df.GetFormat(item.DataFormat);
                }
                temp = listColors.Where(t => t.RGB == item.FillBackgroundColor);
                if (temp != null && temp.Any())
                    cellStyle.FillBackgroundColor = temp.FirstOrDefault().Index;

                temp = listColors.Where(t => t.RGB == item.FillForegroundColor);
                if (temp != null && temp.Any())
                    cellStyle.FillForegroundColor = temp.FirstOrDefault().Index;

                cellStyle.FillPattern = item.FillPattern;
                cellStyle.Indention = item.Indention;
                cellStyle.IsHidden = item.IsHidden;
                cellStyle.IsLocked = item.IsLocked;
                temp = listColors.Where(t => t.RGB == item.LeftBorderColor);
                if (temp != null && temp.Any())
                    cellStyle.LeftBorderColor = temp.FirstOrDefault().Index;

                temp = listColors.Where(t => t.RGB == item.RightBorderColor);
                if (temp != null && temp.Any())
                    cellStyle.RightBorderColor = temp.FirstOrDefault().Index;

                cellStyle.Rotation = item.Rotation;
                cellStyle.ShrinkToFit = item.ShrinkToFit;
                cellStyle.VerticalAlignment = item.VerticalAlignment;
                cellStyle.WrapText = item.WrapText;
                if (item.Font != null)
                {
                    IFont font = workBook.CreateFont();
                    font.Boldweight = item.Font.Boldweight;
                    font.Charset = item.Font.Charset;
                    temp = listColors.Where(t => t.RGB == item.Font.Color);
                    if (temp != null && temp.Any())
                        font.Color = temp.FirstOrDefault().Index;
                    font.FontHeight = item.Font.FontHeight;
                    font.FontHeightInPoints = item.Font.FontHeightInPoints;
                    font.FontName = item.Font.FontName;
                    font.IsItalic = item.Font.IsItalic;
                    font.IsStrikeout = item.Font.IsStrikeout;
                    font.Underline = item.Font.Underline;
                    cellStyle.SetFont(font);
                }

                styleEntity = new CellStyleEntity()
                {
                    CellStyleIndex = item.CellStyleIndex,
                    CellStyle = cellStyle
                };
                listCellStyle.Add(styleEntity);
            }

            foreach (var item in helpEntity.ExportColumns)
            {
                var temp = listCellStyle.Where(t => t.CellStyleIndex == item.CellStyleIndex);
                if (temp != null && temp.Any())
                    item.CellStyle = temp.FirstOrDefault().CellStyle;
            }
            return helpEntity;
        }

        private IEnumerable<ColorEntity> ProcessColor(ExportRunEntity helpEntity)
        {
            if (helpEntity.ExportColors == null || helpEntity.ExportColors.Any() == false)
                return new List<ColorEntity>();
            IEnumerable<string> rgbs = helpEntity.ExportColors.Select(t => t.RGB).Distinct();
            workBook.SetCustomColor(rgbs);

            helpEntity.ExportColors.ToList().ForEach(t =>
            {
                t.Index = workBook.GetCustomColor(t.RGB);
            });
            return helpEntity.ExportColors;
        }

        private void ProcessSheet(ISheet sheet, ExportRunEntity helpEntity)
        {
            ProcessPicture(sheet, helpEntity.ExportPictures);
            ProcessComment(sheet, helpEntity.ExportComments);
        }

        /// <summary>
        /// 处理图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="ExportPictures"></param>
        /// <returns></returns>
        private void ProcessPicture(ISheet sheet, IEnumerable<ExportPictureEntity> ExportPictures)
        {
            if (ExportPictures == null || ExportPictures.Any() == false)
                return;
            foreach (var item in ExportPictures)
            {
                IRow row = sheet.GetRow(item.RowIndex);
                if (row == null)
                    row = sheet.CreateRow(item.RowIndex);
                ICell cell = row.GetCell(item.ColIndex);
                if (cell == null)
                    cell = row.CreateCell(item.ColIndex);
                PictureEntity entity = cell.GetPictureData(item.Url, item.UrlType);
                if (entity == null) continue;
                IPicture pic = sheet.CreateDrawingPatriarch().CreatePicture(entity.Anchor, entity.PictureIndex);
                if (item.Scale != 1)
                    pic.Resize(item.Scale);
            }
        }

        private void ProcessComment(ISheet sheet, IEnumerable<CommentEntity> ExportComments)
        {
            if (ExportComments == null || ExportComments.Any() == false)
                return;
            foreach (var item in ExportComments)
            {
                IRow row = sheet.GetRow(item.RowIndex);
                if (row == null)
                    row = sheet.CreateRow(item.RowIndex);
                ICell cell = row.GetCell(item.ColIndex);
                if (cell == null)
                    cell = row.CreateCell(item.ColIndex);
                cell.SetCellComment(item);
            }
        }

        /// <summary>
        /// 把导出数据写入Http
        /// </summary>
        /// <param name="fileName"></param>
        private void HttpWrite(string fileName)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                workBook.Write(ms);
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearHeaders();
                HttpContext.Current.Response.Buffer = true;
                if (!StringUtil.Contains(HttpContext.Current.Request.UserAgent, "firefox", true) &&
                !StringUtil.Contains(HttpContext.Current.Request.UserAgent, "chrome", true))
                    fileName = StringUtil.UrlEncode(fileName, Encoding.UTF8, false);
                fileName = fileName.Replace("\"", "");
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + fileName);
                // 加入ContentType 防止火狐浏览器导出时直接导出Html，让其默认Excel导出
                HttpContext.Current.Response.ContentType = "application/ms-excel";
                HttpContext.Current.Response.BinaryWrite(ms.ToArray());
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.End();
                ms.Close();
            }
        }

        /// <summary>
        /// 获取合并标题开始列位置
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="mergeName"></param>
        /// <returns></returns>
        private int GetIndex(IEnumerable<ExportColumnEntity> columns, string mergeName)
        {
            int index = 0;
            foreach (var item in columns)
            {
                if (item.MergeName == mergeName)
                    break;
                index++;
            }
            return index;
        }

        /// <summary>
        /// 是否是等差数量
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private bool EqualDiif(List<int> list, int diff)
        {
            bool result = true;
            // 判断是否是等差数列
            for (int i = 0; i < list.Count; i++)
            {
                if (i + 1 < list.Count)
                {
                    int curValue = list[i];
                    int lastValue = list[i + 1];
                    if (lastValue - curValue != diff)
                    {
                        result = false;
                        break;
                    }
                }
            }
            return result;
        }

    }
}
