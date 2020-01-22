using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using UMS.Framework.NpoiUtil.Model;
using UMS.Framework.NpoiUtil.Util;

namespace UMS.Framework.NpoiUtil
{
    public class ImportUtil
    {
        public Tuple<string, IEnumerable<ImportDataEntity>> ProcessWorkBook<T>(string path, string suffix)
            where T : class, new()
        {
            if (string.IsNullOrEmpty(path))
            {
                return new Tuple<string, IEnumerable<ImportDataEntity>>("导入模板路径为空，请联系管理员", null);
            }
            string[] suffixs = { ".xls", ".xlsx" };
            if (suffixs.Contains(suffix) == false)
            {
                return new Tuple<string, IEnumerable<ImportDataEntity>>("导入类型不支持，请联系管理员", null);
            }
            IWorkbook workBook = null;
            ISheet sheet = null;
            string errorMessage = null;
            List<ImportDataEntity> list = new List<ImportDataEntity>();
            ImportDataEntity entity = null;
            using (FileStream fs = new FileStream(HttpContext.Current.Server.MapPath(path), FileMode.Open, FileAccess.Read))
            {
                switch (suffix.ToLower())
                {
                    case ".xls":
                        workBook = new HSSFWorkbook(fs);
                        break;
                    case ".xlsx":
                        workBook = new XSSFWorkbook(fs);
                        break;
                }
                int sheetNum = workBook.NumberOfSheets;
                for (int i = 0; i < sheetNum; i++)
                {
                    sheet = workBook.GetSheetAt(i);
                    entity = ProcessSheetData<T>(workBook, sheet, ref errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                        break;
                    list.Add(entity);
                }
            }
            if (string.IsNullOrEmpty(errorMessage) == false)
                return new Tuple<string, IEnumerable<ImportDataEntity>>(errorMessage, null);
            return new Tuple<string, IEnumerable<ImportDataEntity>>("", list.AsEnumerable());
        }

        private ImportDataEntity ProcessSheetData<T>(IWorkbook workBook, ISheet sheet, ref string errorMessage)
            where T : class, new()
        {
            IEnumerable<CellMarkEntity> cellMarks = ProcessMarkCell<T>(sheet, ref errorMessage);
            if (!string.IsNullOrEmpty(errorMessage))
                return null;
            //IEnumerable<Attribute> dataItems = cellMarks.Where(t => t.Attribute is DataItemAttribute).Select(t => t.Attribute);
            //IEnumerable<Attribute> departs = cellMarks.Where(t => t.Attribute is DepartmentAttribute).Select(t => t.Attribute);
            IEnumerable<CellMergeEntity> cellMerges = ProcessMergeCell(sheet);
            var marks1 = cellMarks.Where(t => t.TransGain == false);
            var marks2 = cellMarks.Where(t => t.TransGain == true);
            var tupe1 = ProcessTransverseData<T>(sheet, marks2, cellMerges);
            var tupe2 = ProcessVerticalData<T>(sheet, marks1, cellMerges);
            return new ImportDataEntity
            {
                ParemtEntity = tupe1.Item1,
                ParentCells = tupe1.Item2,
                ChildEntity = tupe2.Item1,
                ChildCells = tupe2.Item2
            };
        }

        private Tuple<T, IEnumerable<CellDataEntity>> ProcessTransverseData<T>(ISheet sheet, IEnumerable<CellMarkEntity> cellMarks, IEnumerable<CellMergeEntity> cellMerges)
            where T : class, new()
        {
            T entity = new T();
            List<CellDataEntity> cellDatas = new List<CellDataEntity>();
            CellDataEntity cellData = null;
            foreach (var item in cellMarks)
            {
                cellData = ProcessCell(sheet, item, cellMerges);
                Type originalType = item.PropertyInfo.PropertyType.GetGenericArguments()[0];
                if (originalType == typeof(int))
                {
                    item.PropertyInfo.SetValue(entity, int.Parse(cellData.CellValue));
                }
                else if (originalType == typeof(double))
                {
                    item.PropertyInfo.SetValue(entity, double.Parse(cellData.CellValue));
                }
                else if (originalType == typeof(DateTime))
                {
                    item.PropertyInfo.SetValue(entity, DateTime.Parse(cellData.CellValue));
                }else if(originalType == typeof(bool))
                {
                    item.PropertyInfo.SetValue(entity, Boolean.Parse(cellData.CellValue));
                }else
                {
                    item.PropertyInfo.SetValue(entity, cellData.CellValue);
                }
                cellDatas.Add(cellData);
            }
            return new Tuple<T, IEnumerable<CellDataEntity>>(entity, cellDatas);
        }

        private Tuple<IEnumerable<T>, IEnumerable<CellDataEntity>> ProcessVerticalData<T>(ISheet sheet, IEnumerable<CellMarkEntity> cellMarks, IEnumerable<CellMergeEntity> cellMerges)
            where T : class, new()
        {
            int beginRow = cellMarks.Min(t => t.RowIndex) + 1;
            List<T> list = new List<T>();
            List<CellDataEntity> cellDatas = new List<CellDataEntity>();
            CellDataEntity cellData = null;
            T entity = null;
            IRow row = null;
            cellMarks = cellMarks.OrderBy(t => t.CellIndex);
            for (var rowIndex = beginRow; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
            {
                entity = new T();
                row = sheet.GetRow(rowIndex);
                foreach (var item in cellMarks)
                {
                    cellData = ProcessCell(sheet, item, cellMerges, row);
                    if (cellData == null) continue;
                    Type originalType = item.PropertyInfo.PropertyType.GetGenericArguments()[0];
                    if(originalType == typeof(int))
                    {
                        int intVal = 0;
                        cellData.ConvertSuccess = int.TryParse(cellData.CellValue, out intVal);
                        item.PropertyInfo.SetValue(entity, intVal);
                    }
                    else if(originalType == typeof(double))
                    {
                        double dbVal = 0;
                        cellData.ConvertSuccess = double.TryParse(cellData.CellValue, out dbVal);
                        item.PropertyInfo.SetValue(entity, dbVal);
                    }else if(originalType == typeof(DateTime))
                    {
                        DateTime dtVal = DateTime.Now;
                        cellData.ConvertSuccess = DateTime.TryParse(cellData.CellValue, out dtVal);
                        item.PropertyInfo.SetValue(entity, dtVal);
                    }
                    else if (originalType == typeof(bool))
                    {
                        bool blVal = true;
                        cellData.ConvertSuccess = bool.TryParse(cellData.CellValue, out blVal);
                        item.PropertyInfo.SetValue(entity, blVal);
                    }else
                    {
                        item.PropertyInfo.SetValue(entity, cellData.CellValue);
                    }
                    cellDatas.Add(cellData);
                }
                list.Add(entity);
            }
            return new Tuple<IEnumerable<T>, IEnumerable<CellDataEntity>>(list.AsEnumerable(), cellDatas.AsEnumerable());
        }

        private CellDataEntity ProcessCell(ISheet sheet, CellMarkEntity item, IEnumerable<CellMergeEntity> cellMerges, IRow row = null)
        {
            CellDataEntity cellData = new CellDataEntity();
            cellData.TitleName = item.TitleName;
            cellData.ColumnName = item.ColumnName;
            cellData.ColIndex = item.CellIndex;
            ICell cell = null;
            int rowIndex = 0;
            if (row == null)
            {
                rowIndex = item.RowIndex;
                cell = sheet.GetRow(item.RowIndex).GetCell(item.CellIndex + 1);
            }
            else
            {
                rowIndex = row.RowNum;
                cell = row.GetCell(item.CellIndex);
            } 
            if (cell == null)
                return null;
            cellData.RowIndex = rowIndex;
            if (cell.IsMergedCell == false) 
            {
                cellData.CellPostion = string.Concat(ExportExcelUtil.IndexToColName(item.CellIndex), rowIndex + 1);
                Type originalType = item.PropertyInfo.PropertyType.GetGenericArguments()[0];
                string cellValue = "";
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                        {
                            cellValue = cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        }else
                        {
                            cellValue = cell.NumericCellValue.ToString();
                        }
                        break;
                    case CellType.String:
                        cellValue = cell.StringCellValue;
                        break;
                    case CellType.Boolean:
                        cellValue = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Formula:
                        if (originalType == typeof(DateTime))
                        {
                            cellValue = cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        }
                        else
                        {
                            cell.SetCellType(CellType.String);
                            cellValue = cell.ToString();
                        }
                        break;
                }
                cellData.IsEmpty = string.IsNullOrEmpty(cellValue);
                cellData.CellValue = cellValue;
                //if (originalType == typeof(DateTime))
                //{
                //    cellData.CellValue = cell.DateCellValue.ToString();
                //}
                //else
                //{
                //    if (cell.CellType == CellType.Formula)
                //        cell.SetCellType(CellType.String);
                //    cellData.CellValue = cell.ToString();
                //}
            }
            else
            {
                CellDimension dimension = ProcessCellDimension(cell, cellMerges);
                var merges = cellMerges.Where(t => t.FirstRow == dimension.FirstRowIndex && t.FirstColumn == dimension.FirstColIndex);
                cellData.CellPostion = string.Concat(ExportExcelUtil.IndexToColName(dimension.FirstColIndex), dimension.FirstRowIndex + 1);
                if (merges != null && merges.Any())
                {
                    string cellValue = merges.FirstOrDefault().DataCell.ToString();
                    cellData.CellValue = cellValue;
                }
            }
            cellData.Length = StringUtil.StrLength(cellData.CellValue);
            return cellData;
        }

        private IEnumerable<CellMarkEntity> ProcessMarkCell<T>(ISheet sheet, ref string errorMessage)
            where T : class, new()
        {
            IWorkbook workBook = sheet.Workbook;
            int NumberOfNames = workBook.NumberOfNames;
            List<CellMarkEntity> list = new List<CellMarkEntity>();
            CellMarkEntity markEntity = null;
            T entity = new T();
            PropertyInfo[] infos = entity.GetType().GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            for (int index = 0; index < NumberOfNames; index++)
            {
                markEntity = new CellMarkEntity();
                IName iname = workBook.GetNameAt(index);
                string region = iname.RefersToFormula.Replace(iname.SheetName, "").Replace("!", "").Replace("$", "");
                string[] regions = region.Split(':');
                int rowIndex = 0;
                int colIndex = 0;
                // 获取行位置和列位置
                DealPostion(regions[0], ref rowIndex, ref colIndex);
                if (iname.NameName.StartsWith("&"))
                    markEntity.TransGain = true;
                markEntity.CellIndex = colIndex;
                markEntity.RowIndex = rowIndex; 
                markEntity.IName = iname;
                markEntity.TitleName = sheet.GetRow(rowIndex).GetCell(colIndex).ToString();
                string name = iname.NameName.Replace("$", "").Replace("&", "");
                IEnumerable<PropertyInfo> tempInfo = infos.Where(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (tempInfo != null && tempInfo.Any())
                {
                    PropertyInfo info = tempInfo.FirstOrDefault();
                    markEntity.ColumnName = info.Name;
                    markEntity.PropertyInfo = info;
                    IEnumerable<Attribute> attributes = info.GetCustomAttributes();
                    if (attributes != null && attributes.Any())
                        markEntity.Attribute = attributes.FirstOrDefault();
                    list.Add(markEntity);
                }
                else
                {
                    errorMessage = "第【{0}】行第【{1}】列配置不正确，请联系管理员";
                    errorMessage = string.Format(errorMessage, rowIndex + 1, ExportExcelUtil.IndexToColName(colIndex));
                    break;
                }
            }
            return list.AsEnumerable();
        }

        private IEnumerable<CellMergeEntity> ProcessMergeCell(ISheet sheet)
        {
            int NumMergedRegions = sheet.NumMergedRegions;
            List<CellMergeEntity> list = new List<CellMergeEntity>();
            CellMergeEntity entity = null;
            for (var index = 0; index < NumMergedRegions; index++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(index);
                entity = new CellMergeEntity();
                entity.FirstColumn = range.FirstColumn;
                entity.FirstRow = range.FirstRow;
                entity.LastColumn = range.LastColumn;
                entity.LastRow = range.LastRow;
                entity.DataCell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                list.Add(entity);
            }
            return list;
        }

        private CellDimension ProcessCellDimension(ICell cell, IEnumerable<CellMergeEntity> cellMerges)
        {
            CellDimension dimension = new CellDimension();

            var merges = cellMerges.Where(t => t.FirstRow <= cell.RowIndex && t.LastRow >= cell.RowIndex && t.FirstColumn <= cell.ColumnIndex && t.LastColumn >= cell.ColumnIndex);
            if (merges != null && merges.Any())
            {
                var merge = merges.FirstOrDefault();
                dimension = new CellDimension
                {
                    Cell = cell,
                    RowSpan = merge.LastRow - merge.FirstRow + 1,
                    ColSpan = merge.LastColumn - merge.FirstColumn + 1,
                    FirstRowIndex = merge.FirstRow,
                    LastRowIndex = merge.LastRow,
                    FirstColIndex = merge.FirstColumn,
                    LastColIndex = merge.LastColumn
                };
            }
            return dimension;
        }

        /// <summary>
        /// 处理单元格区域
        /// </summary>
        /// <param name="postion"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        private void DealPostion(string postion, ref int rowIndex, ref int colIndex)
        {
            if (postion.Length < 2)
                throw new Exception("invalid parameter");
            colIndex = ExportExcelUtil.ColNameToIndex(postion.Substring(0, 1));
            rowIndex = int.Parse(postion.Substring(1, postion.Length - 1)) - 1;
        }

    }
}
