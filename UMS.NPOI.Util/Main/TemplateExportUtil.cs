using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using UMS.Framework.NpoiUtil.Model;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil
{
    public class TemplateExportUtil
    {
        IWorkbook workBook = null;
        public string RunExport(ExportTemplateEntity helpEntity, params DynamicEntity[] entitys)
        {
            string errorMessage = string.Empty;
            IEnumerable<CellHelpEntity> list = null;
            try
            {
                list = BeforeBuild(helpEntity);
                if (list == null || list.Any() == false)
                    return errorMessage = "配置模板为空";
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            if (!string.IsNullOrEmpty(errorMessage)) return errorMessage;

            switch (helpEntity.TemplateType)
            {
                case TemplateType.Table:
                    BuildTableWorkbook(helpEntity, list, entitys);
                    break;
                case TemplateType.Single:
                    BuildSingleWorkbook(helpEntity, list, entitys);
                    break;
            }
            workBook.HttpWrite(string.Concat(helpEntity.FileName, ".", helpEntity.Suffix));
            return errorMessage;
        }

        public string RunExport(ExportTemplateEntity helpEntity, params Tuple<DynamicEntity, IEnumerable<DynamicEntity>>[] tuples)
        {
            string errorMessage = string.Empty;
            IEnumerable<CellHelpEntity> list = null;
            try
            {
                list = BeforeBuild(helpEntity);
                if (list == null || list.Any() == false)
                    return errorMessage = "配置模板为空";
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            if (!string.IsNullOrEmpty(errorMessage)) return errorMessage;
            switch (helpEntity.TemplateType)
            {
                case TemplateType.Bill:
                    BuildBillWorkbook(helpEntity, list, tuples);
                    break;
            }
            workBook.HttpWrite(string.Concat(helpEntity.FileName, ".", helpEntity.Suffix));
            return errorMessage;
        }

        /// <summary>
        /// 生成工作簿
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="dtData"></param>
        /// <returns></returns>
        private string BuildTableWorkbook(ExportTemplateEntity helpEntity, IEnumerable<CellHelpEntity> list, params DynamicEntity[] entitys)
        {
            string errorMessage = string.Empty;

            List<ISheet> listSheet = new List<ISheet>();
            ISheet sheet = list.FirstOrDefault().Cell.Sheet;
            int sheetNum = workBook.GetSheetIndex(sheet.SheetName);
            listSheet.Add(sheet);
            int len = entitys.Length;
            if (len > 1)
            {
                for (var i = 0; i < len - 1; i++)
                {
                    ISheet tempSheet = workBook.CloneSheet(sheetNum);
                    listSheet.Add(tempSheet);
                }
            }
            int index = 0;
            foreach (var entity in entitys)
            {
                BuildTableSheet(listSheet[index], list, entity);
                index++;
            }
            return errorMessage;
        }

        private string BuildSingleWorkbook(ExportTemplateEntity helpEntity, IEnumerable<CellHelpEntity> list, IEnumerable<DynamicEntity> entitys)
        {
            string errorMessage = string.Empty;
            ISheet sheet = list.FirstOrDefault().Cell.Sheet;

            IRow row = null;
            ICell cell = null;
            int index = 0;
            foreach (var entity in entitys)
            {
                foreach (var dimension in list)
                {
                    row = sheet.GetRow(dimension.FirstRowIndex + index);
                    if (row == null) row = sheet.CreateRow(dimension.FirstRowIndex + index);
                    cell = row.GetCell(dimension.FirstColIndex);
                    if (cell == null) cell = row.CreateCell(dimension.FirstColIndex);
                    var value = entity.GetPropertyValue(dimension.Name).ToString();
                    cell.SetCellValue(value);
                }
                index++;
            }
            return errorMessage;
        }

        private string BuildBillWorkbook(ExportTemplateEntity helpEntity, IEnumerable<CellHelpEntity> list, params Tuple<DynamicEntity, IEnumerable<DynamicEntity>>[] tuples)
        {
            string errorMessage = string.Empty;

            List<ISheet> listSheet = new List<ISheet>();
            ISheet sheet = list.FirstOrDefault().Cell.Sheet;
            int sheetNum = workBook.GetSheetIndex(sheet.SheetName);
            listSheet.Add(sheet);
            int len = tuples.Length;
            if (len > 1)
            {
                for (var i = 0; i < len - 1; i++)
                {
                    ISheet tempSheet = workBook.CloneSheet(sheetNum);
                    listSheet.Add(tempSheet);
                }
            }
            int index = 0;
            foreach (var tupe in tuples)
            {
                BuildBillSheet(listSheet[index], list, tupe.Item1, tupe.Item2);
                index++;
            }
            return errorMessage;
        }

        private void BuildTableSheet(ISheet sheet, IEnumerable<CellHelpEntity> list, DynamicEntity entity)
        {
            dynamic o = new ExpandoObject();
            ICell cell = null;
            foreach (var dimension in list)
            {
                cell = GetCell(sheet, dimension);
                var value = entity.GetPropertyValue(dimension.Name).ToString();
                cell.SetCellValue(value);
                //cell.CellStyle = dimension.Cell.CellStyle;
                //CellRangeAddress address = new CellRangeAddress(dimension.FirstRowIndex, dimension.LastRowIndex, dimension.FirstColIndex, dimension.LastColIndex);
                //sheet.AddMergedRegion(address);
            }
        }

        private void BuildBillSheet(ISheet sheet, IEnumerable<CellHelpEntity> list, DynamicEntity main, IEnumerable<DynamicEntity> subs)
        {
            ICell cell = null;
            var listMain = list.Where(t => t.Type == ConfigType.One);
            var listSub = list.Where(t => t.Type == ConfigType.More);
            foreach (var dm in listMain)
            {
                cell = GetCell(sheet, dm);
                var value = main.GetPropertyValue(dm.Name, false);
                if (value == null)
                    cell.SetCellValue(string.Empty);
                else
                    cell.SetCellValue(value.ToString());
            }
            int index = 0;
            foreach (var sub in subs)
            {
                foreach (var ds in listSub)
                {
                    cell = GetCell(sheet, ds, index, 0, true);
                    var value = sub.GetPropertyValue(ds.Name,false);
                    if (value == null)
                        cell.SetCellValue(string.Empty);
                    else
                        cell.SetCellValue(value.ToString());
                }
                index++;
            }
            int num = listSub.Max(t => t.LastRowIndex);
            var mainDown = listMain.Where(t => t.FirstRowIndex >= num);
            foreach (var dm in mainDown)
            {
                cell = GetCell(sheet, dm);
                var value = main.GetPropertyValue(dm.Name, false);
                if (value == null)
                    cell.SetCellValue(string.Empty);
                else
                    cell.SetCellValue(value.ToString());
            }
        }

        private ICell GetCell(ISheet sheet, CellDimension cd, int rowNum = 0, int colNum = 0, bool sub = false)
        {
            IRow row = null;
            ICell cell = null;
            if (!sub)
            {
                row = sheet.GetRow(cd.FirstRowIndex + rowNum);
            }
            else
            {
                row = sheet.GetRow(cd.FirstRowIndex + rowNum);
            }
            if (row == null)
            {
                row = sheet.CreateRow(cd.FirstRowIndex + rowNum);
                row.RowStyle = sheet.GetRow(cd.FirstRowIndex).RowStyle;
            }
            cell = row.GetCell(cd.FirstColIndex + colNum);
            if (cell == null)
            {
                cell = row.CreateCell(cd.FirstColIndex + colNum);
                ICell oldCell = sheet.GetRow(cd.FirstRowIndex).GetCell(cd.FirstColIndex);
                //cell.CellStyle = sheet.GetRow(cd.FirstRowIndex).GetCell(cd.FirstColIndex).CellStyle;
                cell.CellStyle = oldCell.CellStyle;
            }
            if (cd.IsMergeCell)
            {
                CellRangeAddress address = new CellRangeAddress(cd.FirstRowIndex + rowNum, cd.FirstRowIndex + rowNum, cd.FirstColIndex, cd.LastColIndex);
                sheet.AddMergedRegion(address);
                for (var index = cd.FirstColIndex + 1; index <= cd.LastColIndex; index++)
                {
                    ICell cellTemp = sheet.GetRow(cd.FirstRowIndex).GetCell(index);
                    row.CreateCell(index).CellStyle = cellTemp.CellStyle;
                }
            }
            return cell;
        }

        #region 基础处理

        private IEnumerable<CellHelpEntity> BeforeBuild(ExportTemplateEntity helpEntity)
        {
            string path = helpEntity.Path;
            switch (helpEntity.PathType)
            {
                case PathType.Http:
                    //path = HttpContext.Current.Server.MapPath(path);
                    byte[] bytes = GetHttpUrlData(path);
                    using (MemoryStream ms = new MemoryStream(bytes, 0, bytes.Length))
                    {
                        switch (helpEntity.Suffix)
                        {
                            case ExportExcelSuffix.xls:
                                workBook = new HSSFWorkbook(ms);
                                break;
                            case ExportExcelSuffix.xlsx:
                                workBook = new XSSFWorkbook(ms);
                                break;
                        }
                    }
                    break;
                case PathType.File:
                    path = helpEntity.Path;
                    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        switch (helpEntity.Suffix)
                        {
                            case ExportExcelSuffix.xls:
                                workBook = new HSSFWorkbook(fs);
                                break;
                            case ExportExcelSuffix.xlsx:
                                workBook = new XSSFWorkbook(fs);
                                break;
                        }
                    }
                    break;
            }
            ISheet sheet = workBook.GetSheetAt(0);
            List<CellHelpEntity> list = new List<CellHelpEntity>();
            CellHelpEntity entity = null;
            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                for (short colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell.IsMergedCell && string.IsNullOrEmpty(cell.ToString())) continue;
                    CellDimension dimension = cell.GetSpan();
                    entity = ProcessCell(dimension);
                    entity.SorceRow = row;
                    if (entity != null)
                        list.Add(entity);
                }
            }
            return list;
        }

        private byte[] GetHttpUrlData(string url)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Method = "GET";
            byte[] buffurPic = null;
            try
            {
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
                StreamReader reader = new StreamReader(webResponse.GetResponseStream(), Encoding.UTF8);
                Stream stream = webResponse.GetResponseStream();
                MemoryStream ms = null;
                Byte[] buffer = new Byte[webResponse.ContentLength];
                int offset = 0, actuallyRead = 0;
                do
                {
                    actuallyRead = stream.Read(buffer, offset, buffer.Length - offset);
                    offset += actuallyRead;
                }
                while (actuallyRead > 0);
                ms = new MemoryStream(buffer);

                buffurPic = ms.ToArray();
            }
            catch { }

            return buffurPic;
        }

        private CellHelpEntity ProcessCell(CellDimension entity)
        {
            CellHelpEntity helpEntity = null;
            var value = entity.Cell.ToString();
            //if (!value.StartsWith("$"))
            //    return null;
            helpEntity = new CellHelpEntity(entity);
            //value = value.Replace(" ", "").Replace("&=", "");
            value = value.Trim();
            var temp = value.Substring(0, 1);
            helpEntity.Name = value.Replace(temp, "");
            switch (temp)
            {
                case "$":
                    helpEntity.Type = ConfigType.One;
                    break;
                case "&":
                    helpEntity.Type = ConfigType.More;
                    break;
                default:
                    helpEntity = null;
                    break;
            }
            return helpEntity;
        }

        private ISheet CreateNewSheet(IWorkbook workBook, ISheet oldSheet, string sheetName = "")
        {
            ISheet sheet = null;
            if (string.IsNullOrEmpty(sheetName))
                sheet = workBook.CreateSheet();
            else
                sheet = workBook.CreateSheet(sheetName);
            sheet.Autobreaks = oldSheet.Autobreaks;
            sheet.DefaultColumnWidth = oldSheet.DefaultColumnWidth;
            sheet.DefaultRowHeight = oldSheet.DefaultRowHeight;
            sheet.DefaultRowHeightInPoints = oldSheet.DefaultRowHeightInPoints;
            sheet.DisplayFormulas = oldSheet.DisplayFormulas;
            sheet.DisplayGridlines = oldSheet.DisplayGridlines;
            sheet.DisplayGuts = oldSheet.DisplayGuts;
            sheet.DisplayRowColHeadings = oldSheet.DisplayRowColHeadings;
            sheet.DisplayZeros = oldSheet.DisplayZeros;
            sheet.FitToPage = oldSheet.FitToPage;
            sheet.ForceFormulaRecalculation = oldSheet.ForceFormulaRecalculation;
            sheet.HorizontallyCenter = sheet.HorizontallyCenter;
            sheet.IsPrintGridlines = oldSheet.IsPrintGridlines;
            sheet.IsRightToLeft = oldSheet.IsRightToLeft;
            sheet.IsSelected = oldSheet.IsSelected;
            //sheet.LeftCol = oldSheet.LeftCol;
            sheet.RepeatingColumns = oldSheet.RepeatingColumns;
            sheet.RepeatingRows = oldSheet.RepeatingRows;
            sheet.RowSumsBelow = oldSheet.RowSumsBelow;
            sheet.RowSumsRight = oldSheet.RowSumsRight;
            //sheet.TabColorIndex = oldSheet.TabColorIndex;
            //sheet.TopRow = oldSheet.TopRow;
            sheet.VerticallyCenter = oldSheet.VerticallyCenter;
            return sheet;
        }

        #endregion
    }
}
