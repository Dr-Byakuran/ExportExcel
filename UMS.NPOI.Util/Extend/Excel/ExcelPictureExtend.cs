using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// <summary>
    /// 图表处理扩展
    /// </summary>
    public static class ExcelPictureExtend
    {
        #region 图片信息获取

        /// <summary>
        /// 获取重定义名称实体集合
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public static IEnumerable<INameInfo> GetAllINameInfo(this IWorkbook workBook)
        {
            List<INameInfo> list = new List<INameInfo>();
            for (int index = 0; index < workBook.NumberOfNames; index++)
            {
                IName iname = workBook.GetNameAt(index);
                if (iname.IsDeleted || iname.IsFunctionName)
                    continue;
                INameInfo info = DealIName(iname);
                if (info == null)
                    continue;
                list.Add(info);
            }

            return list;
        }

        /// <summary>
        /// 获取所有图片信息
        /// 参考地址：https://www.cnblogs.com/hanzhaoxin/p/4442369.html
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IEnumerable<PictureInfo> GetAllPictureInfo(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }
        #endregion

        #region 处理图片获取

        /// <summary>
        /// 获取所有图片信息
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static List<PictureInfo> GetAllPictureInfos(this ISheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet)sheet, firstRow, lastRow, firstCol, lastCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet)sheet, firstRow, lastRow, firstCol, lastCol, onlyInternal);
            }
            else
            {
                throw new Exception("Unhandled type, Not added for this type: GetAllPicturesInfos() Extension method.");
            }
        }

        /// <summary>
        /// 获取所有图片信息：低版本
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal)
        {
            List<PictureInfo> picturesInfoList = new List<PictureInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture)shape;
                        var anchor = (HSSFClientAnchor)shape.Anchor;

                        if (IsInternalOrIntersect(firstRow, lastRow, firstCol, lastCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                        {
                            picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.SuggestFileExtension(), picture.PictureData.Data));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        /// <summary>
        /// 获取所有图片信息：高版本
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static List<PictureInfo> GetAllPictureInfos(XSSFSheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool onlyInternal)
        {
            List<PictureInfo> picturesInfoList = new List<PictureInfo>();

            var documentPartList = sheet.GetRelations();
            foreach (var documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing)documentPart;
                    var shapeList = drawing.GetShapes();
                    foreach (var shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture)shape;
                            picture.PictureData.SuggestFileExtension();
                            var anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(firstRow, lastRow, firstCol, lastCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.SuggestFileExtension(), picture.PictureData.Data));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        /// <summary>
        /// 判断
        /// </summary>
        /// <param name="rangeMinRow"></param>
        /// <param name="rangeMaxRow"></param>
        /// <param name="rangeMinCol"></param>
        /// <param name="rangeMaxCol"></param>
        /// <param name="pictureMinRow"></param>
        /// <param name="pictureMaxRow"></param>
        /// <param name="pictureMinCol"></param>
        /// <param name="pictureMaxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }

        #endregion

        #region 重定义名称

        /// <summary>
        /// 处理位置
        /// 当前只处理单个单元格、或连续单元格
        /// 例如：Sheet2!$G$11 或 Sheet2!$D$5:$E$14 这两种情况
        /// </summary>
        /// <param name="iName"></param>
        /// <returns></returns>
        private static INameInfo DealIName(this IName iname)
        {
            if (string.IsNullOrEmpty(iname.RefersToFormula))
                return null;
            // 如果是跨区域则返回空； 例如：Sheet2!$D$5:$E$14,Sheet2!$G$11,Sheet2!$G$10
            string[] regions = iname.RefersToFormula.Split(',');
            if (regions.Length > 1)
                return null;
            INameInfo info = new INameInfo(iname);
            // 先替换不需要字符
            string region = iname.RefersToFormula.Replace(iname.SheetName, "").Replace("!", "").Replace("$", "");
            string[] postions = region.Split(':');

            int rowBegin = 0;
            int rowEnd = 0;
            int colBegin = 0;
            int colEnd = 0;
            if (postions.Length == 1) //单个单元格
            {
                // 处理位置
                DealPostion(postions[0], ref rowBegin, ref colBegin);
                // 位置
                info.FirstRow = info.LastRow = rowBegin;
                info.FirstCol = info.LastCol = colBegin;
            }
            else if (postions.Length == 2)//区域单元格
            {
                // 处理位置
                DealPostion(postions[0], ref rowBegin, ref colBegin);
                DealPostion(postions[1], ref rowEnd, ref colEnd);
                //位置赋值
                info.FirstRow = rowBegin;
                info.LastRow = rowEnd;
                info.FirstCol = colBegin;
                info.LastCol = colEnd;
            }
            if (info.FirstRow == info.LastRow)
                info.EqualRow = true;
            else
                info.EqualRow = false;

            if (info.FirstCol == info.LastCol)
                info.EqualCol = true;
            else
                info.EqualCol = false;
            return info;
        }

        /// <summary>
        /// 对区域进行处理，返回行位置和列位置
        /// </summary>
        /// <param name="postion"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        private static void DealPostion(string postion, ref int rowIndex, ref int colIndex)
        {
            if (postion.Length < 2)
                throw new Exception("invalid parameter");
            colIndex = ExcelExtend.ColumnNameToIndex(postion.Substring(0, 1));
            rowIndex = int.Parse(postion.Substring(1, 1)) - 1;
        }
        #endregion
    }
}
