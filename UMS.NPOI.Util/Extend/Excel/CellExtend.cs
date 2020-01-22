using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UMS.Framework.NpoiUtil.Model;
using UMS.Framework.NpoiUtil.Util;

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
    /// 设置值扩展
    /// </summary>
    public static class CellExtend
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

        /// <summary>
        /// 设置值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="colEntity"></param>
        /// <param name="cellStyle"></param>
        public static void SetCellValue(this ICell cell, object value, ExportColumnEntity colEntity, ICellStyle cellStyle)
        {
            IWorkbook workBook = cell.Sheet.Workbook;
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                cell.CellStyle = cellStyle;
                return;
            }
            string strValue = value.ToString();
            switch (colEntity.CellType)
            {
                case CellType.Blank:
                    cell.SetCellValue(strValue);
                    break;
                case CellType.Boolean:
                    bool blValue = false;
                    bool.TryParse(strValue, out blValue);
                    cell.SetCellValue(blValue);
                    break;
                case CellType.Error:
                    cell.SetCellValue(strValue);
                    break;
                case CellType.Formula:
                    cell.SetCellFormula(strValue);
                    break;
                case CellType.Numeric:
                    double dbValue = 0;
                    double.TryParse(strValue, out dbValue);
                    cell.SetCellValue(dbValue);
                    break;
                case CellType.String:
                    cell.SetCellValue(strValue);
                    break;
                case CellType.Unknown:
                    cell.SetCellValue(strValue);
                    break;
                default:
                    cell.SetCellValue(strValue);
                    break;
            }
            cellStyle.Alignment = colEntity.HAlign;
            cellStyle.VerticalAlignment = colEntity.VAlign;
            if (!string.IsNullOrEmpty(colEntity.DataFormat))
            {
                IDataFormat df = workBook.CreateDataFormat();
                cellStyle.DataFormat = df.GetFormat(colEntity.DataFormat);
            }
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置批注
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="suffix"></param>
        /// <param name="comment"></param>
        /// <param name="col1"></param>
        /// <param name="row1"></param>
        /// <param name="col2"></param>
        /// <param name="row2"></param>
        public static void SetCellComment(this ICell cell, CommentEntity entitiy)
        {
            ISheet sheet = cell.Sheet;
            IClientAnchor clientAnchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            clientAnchor.AnchorType = AnchorType.MoveDontResize.GetType().ToInt();
            clientAnchor.Dx1 = entitiy.Dx1;
            clientAnchor.Dy1 = entitiy.Dy1;
            clientAnchor.Dx2 = entitiy.Dx2;
            clientAnchor.Dy2 = entitiy.Dy2;
            clientAnchor.Col1 = cell.ColumnIndex;
            clientAnchor.Row1 = cell.RowIndex;
            clientAnchor.Col2 = cell.ColumnIndex + entitiy.Width;
            clientAnchor.Row2 = cell.RowIndex + entitiy.Height;

            IDrawing draw = sheet.CreateDrawingPatriarch();
            IComment comment = draw.CreateCellComment(clientAnchor);
            comment.Visible = false;
            if (sheet.Workbook is HSSFWorkbook)
                comment.String = new HSSFRichTextString(entitiy.Text);
            else
                comment.String = new XSSFRichTextString(entitiy.Text);
            cell.CellComment = comment;
        }

        /// <summary>
        /// 获取图片信息
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="url"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static PictureEntity GetPictureData(this ICell cell, string url, UrlType type = UrlType.Base64)
        {
            byte[] data = null;
            switch (type)
            {
                case UrlType.Base64:
                    string[] urls = url.Split(',');
                    // 获取图表
                    data = Convert.FromBase64String(urls[urls.Length - 1]);
                    break;
                case UrlType.Http:
                    data = ExportExcelUtil.GetHttpFile(url);
                    break;
                default:
                    break;
            }
            if (data == null || data.Length == 0)
                return null;
            double scalx = 0;//x轴缩放比例
            double scaly = 0;//y轴缩放比例
            int Dx1 = 0;//图片左边相对excel格的位置(x偏移) 范围值为:0~1023,超过1023就到右侧相邻的单元格里了
            int Dy1 = 0;//图片上方相对excel格的位置(y偏移) 范围值为:0~256,超过256就到下方的单元格里了
            bool bOriginalSize = false;//是否显示图片原始大小 true表示图片显示原始大小  false表示显示图片缩放后的大小
            ///计算单元格的长度和宽度
            double CellWidth = 0;
            double CellHeight = 0;
            int RowSpanCount = cell.GetSpan().RowSpan;//合并的单元格行数
            int ColSpanCount = cell.GetSpan().ColSpan;//合并的单元格列数 
            int j = 0;
            for (j = 0; j < RowSpanCount; j++)//根据合并的行数计算出高度
            {
                CellHeight += cell.Sheet.GetRow(cell.RowIndex + j).Height;
            }
            for (j = 0; j < ColSpanCount; j++)
            {
                CellWidth += cell.Row.Sheet.GetColumnWidth(cell.ColumnIndex + j);
            }
            //单元格长度和宽度与图片的长宽单位互换是根据实例得出
            CellWidth = CellWidth / 35;
            CellHeight = CellHeight / 15;
            ///计算图片的长度和宽度
            MemoryStream ms = new MemoryStream(data);
            Image Img = Bitmap.FromStream(ms, true);
            double ImageOriginalWidth = Img.Width;//原始图片的长度
            double ImageOriginalHeight = Img.Height;//原始图片的宽度
            double ImageScalWidth = 0;//缩放后显示在单元格上的图片长度
            double ImageScalHeight = 0;//缩放后显示在单元格上的图片宽度
            if (CellWidth > ImageOriginalWidth && CellHeight > ImageOriginalHeight)//单元格的长度和宽度比图片的大，说明单元格能放下整张图片，不缩放
            {
                ImageScalWidth = ImageOriginalWidth;
                ImageScalHeight = ImageOriginalHeight;
                bOriginalSize = true;
            }
            else//需要缩放，根据单元格和图片的长宽计算缩放比例
            {
                bOriginalSize = false;
                if (ImageOriginalWidth > CellWidth && ImageOriginalHeight > CellHeight)//图片的长和宽都比单元格的大的情况
                {
                    double WidthSub = ImageOriginalWidth - CellWidth;//图片长与单元格长的差距
                    double HeightSub = ImageOriginalHeight - CellHeight;//图片宽与单元格宽的差距
                    if (WidthSub > HeightSub)//长的差距比宽的差距大时,长度x轴的缩放比为1，表示长度就用单元格的长度大小，宽度y轴的缩放比例需要根据x轴的比例来计算
                    {
                        scalx = 1;
                        scaly = (CellWidth / ImageOriginalWidth) * ImageOriginalHeight / CellHeight;//计算y轴的缩放比例,CellWidth / ImageWidth计算出图片整体的缩放比例,然后 * ImageHeight计算出单元格应该显示的图片高度,然后/ CellHeight就是高度的缩放比例
                    }
                    else
                    {
                        scaly = 1;
                        scalx = (CellHeight / ImageOriginalHeight) * ImageOriginalWidth / CellWidth;
                    }
                }
                else if (ImageOriginalWidth > CellWidth && ImageOriginalHeight < CellHeight)//图片长度大于单元格长度但图片高度小于单元格高度，此时长度不需要缩放，直接取单元格的，因此scalx=1，但图片高度需要等比缩放
                {
                    scalx = 1;
                    scaly = (CellWidth / ImageOriginalWidth) * ImageOriginalHeight / CellHeight;
                }
                else if (ImageOriginalWidth < CellWidth && ImageOriginalHeight > CellHeight)//图片长度小于单元格长度但图片高度大于单元格高度，此时单元格高度直接取单元格的，scaly = 1,长度需要等比缩放
                {
                    scaly = 1;
                    scalx = (CellHeight / ImageOriginalHeight) * ImageOriginalWidth / CellWidth;
                }
                ImageScalWidth = scalx * CellWidth;
                ImageScalHeight = scaly * CellHeight;
            }
            Dx1 = Convert.ToInt32((CellWidth - ImageScalWidth) / CellWidth * 1023 / 2);
            Dy1 = Convert.ToInt32((CellHeight - ImageScalHeight) / CellHeight * 256 / 2);
            int pictureIdx = cell.Sheet.Workbook.AddPicture((Byte[])data, PictureType.PNG);
            IClientAnchor anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.AnchorType = AnchorType.MoveDontResize.GetType().ToInt();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = cell.ColumnIndex + cell.GetSpan().ColSpan;
            anchor.Row1 = cell.RowIndex;
            anchor.Row2 = cell.RowIndex + cell.GetSpan().RowSpan;
            anchor.Dy1 = Dy1;//图片下移量
            anchor.Dx1 = Dx1;//图片右移量，通过图片下移和右移，使得图片能居中显示，因为图片不同文字，图片是浮在单元格上的，文字是钳在单元格里的

            PictureEntity entity = new PictureEntity()
            {
                ScaleX = scalx,
                ScaleY = scaly,
                Anchor = anchor,
                PictureIndex = pictureIdx,
                OriginalSize = bOriginalSize
            };
            return entity;
        }

        /// <summary>
        /// 获取合并信息
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static CellDimension GetSpan(this ICell cell)
        {
            CellDimension cellDimension = new CellDimension()
            {
                Cell = null,
                RowSpan = 1,
                ColSpan = 1,
                FirstRowIndex = cell.RowIndex,
                LastRowIndex = cell.RowIndex,
                FirstColIndex = cell.ColumnIndex,
                LastColIndex = cell.ColumnIndex,
                IsMergeCell = false,
            };
            cell.IsMergeCell(out cellDimension);
            return cellDimension;
        }

        #region 判断是否合并 单元格，并返回合并单元格信息

        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的维度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="dimension">单元格维度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ICell cell, out CellDimension dimension)
        {
            return cell.Sheet.IsMergeCell(cell.RowIndex, cell.ColumnIndex, out dimension);
        }

        /// <summary>
        /// 判断指定行列所在的单元格是否为合并单元格，并且输出该单元格的行列跨度
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="rowSpan">行跨度，返回值最小为1，同时表示没有行合并</param>
        /// <param name="ColSpan">列跨度，返回值最小为1，同时表示没有列合并</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ISheet sheet, int rowIndex, int columnIndex, out int rowSpan, out int ColSpan)
        {
            CellDimension dimension;
            bool result = sheet.IsMergeCell(rowIndex, columnIndex, out dimension);

            rowSpan = dimension.RowSpan;
            ColSpan = dimension.ColSpan;

            return result;
        }

        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的行列跨度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="rowSpan">行跨度，返回值最小为1，同时表示没有行合并</param>
        /// <param name="ColSpan">列跨度，返回值最小为1，同时表示没有列合并</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ICell cell, out int rowSpan, out int ColSpan)
        {
            return cell.Sheet.IsMergeCell(cell.RowIndex, cell.ColumnIndex, out rowSpan, out ColSpan);
        }

        /// <summary>
        /// 返回上一个跨度行，如果rowIndex为第一行，则返回null
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回上一个跨度行</returns>
        public static IRow PrevSpanRow(this ISheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.FuncSheet(rowIndex, columnIndex, (currentDimension, isMerge) =>
            {
                //上一个单元格维度
                CellDimension prevDimension;
                sheet.IsMergeCell(currentDimension.FirstRowIndex - 1, columnIndex, out prevDimension);
                return prevDimension.Cell.Row;
            });
        }

        /// <summary>
        /// 返回下一个跨度行，如果rowIndex为最后一行，则返回null
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回下一个跨度行</returns>
        public static IRow NextSpanRow(this ISheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.FuncSheet(rowIndex, columnIndex, (currentDimension, isMerge) =>
                isMerge ? sheet.GetRow(currentDimension.FirstRowIndex + currentDimension.RowSpan) : sheet.GetRow(rowIndex));
        }

        /// <summary>
        /// 返回上一个跨度行，如果row为第一行，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>返回上一个跨度行</returns>
        public static IRow PrevSpanRow(this IRow row)
        {
            return row.Sheet.PrevSpanRow(row.RowNum, row.FirstCellNum);
        }

        /// <summary>
        /// 返回下一个跨度行，如果row为最后一行，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>返回下一个跨度行</returns>
        public static IRow NextSpanRow(this IRow row)
        {
            return row.Sheet.NextSpanRow(row.RowNum, row.FirstCellNum);
        }

        /// <summary>
        /// 返回上一个跨度列，如果columnIndex为第一列，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回上一个跨度列</returns>
        public static ICell PrevSpanCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) =>
            {
                //上一个单元格维度
                CellDimension prevDimension;
                row.Sheet.IsMergeCell(row.RowNum, currentDimension.FirstColIndex - 1, out prevDimension);
                return prevDimension.Cell;
            });
        }

        /// <summary>
        /// 返回下一个跨度列，如果columnIndex为最后一列，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回下一个跨度列</returns>
        public static ICell NextSpanCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) =>
                row.GetCell(currentDimension.FirstColIndex + currentDimension.ColSpan));
        }

        /// <summary>
        /// 返回上一个跨度列，如果cell为第一列，则返回null
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>返回上一个跨度列</returns>
        public static ICell PrevSpanCell(this ICell cell)
        {
            return cell.Row.PrevSpanCell(cell.ColumnIndex);
        }

        /// <summary>
        /// 返回下一个跨度列，如果columnIndex为最后一列，则返回null
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>返回下一个跨度列</returns>
        public static ICell NextSpanCell(this ICell cell)
        {
            return cell.Row.NextSpanCell(cell.ColumnIndex);
        }

        /// <summary>
        /// 返回指定行索引所在的合并单元格(区域)中的第一行(通常是含有数据的行)
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <returns>返回指定列索引所在的合并单元格(区域)中的第一行</returns>
        public static IRow GetDataRow(this ISheet sheet, int rowIndex)
        {
            return sheet.FuncSheet(rowIndex, 0, (currentDimension, isMerge) => sheet.GetRow(currentDimension.FirstRowIndex));
        }

        /// <summary>
        /// 返回指定列索引所在的合并单元格(区域)中的第一行第一列(通常是含有数据的单元格)
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>返回指定列索引所在的合并单元格(区域)中的第一行第一列</returns>
        public static ICell GetCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) => currentDimension.Cell);
        }

        private static T FuncSheet<T>(this ISheet sheet, int rowIndex, int columnIndex, Func<CellDimension, bool, T> func)
        {
            //当前单元格维度
            CellDimension currentDimension;
            //是否为合并单元格
            bool isMerge = sheet.IsMergeCell(rowIndex, columnIndex, out currentDimension);

            return func(currentDimension, isMerge);
        }

        /// <summary>
        /// 判断指定行列所在的单元格是否为合并单元格，并且输出该单元格的维度
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="dimension">单元格维度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ISheet sheet, int rowIndex, int columnIndex, out CellDimension dimension)
        {
            dimension = new CellDimension
            {
                Cell = null,
                RowSpan = 1,
                ColSpan = 1,
                FirstRowIndex = rowIndex,
                LastRowIndex = rowIndex,
                FirstColIndex = columnIndex,
                LastColIndex = columnIndex,
                IsMergeCell = false
            };

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                sheet.IsMergedRegion(range);
                if ((rowIndex >= range.FirstRow && range.LastRow >= rowIndex) && (columnIndex >= range.FirstColumn && range.LastColumn >= columnIndex))
                {
                    dimension.Cell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                    dimension.RowSpan = range.LastRow - range.FirstRow + 1;
                    dimension.ColSpan = range.LastColumn - range.FirstColumn + 1;
                    dimension.FirstRowIndex = range.FirstRow;
                    dimension.LastRowIndex = range.LastRow;
                    dimension.FirstColIndex = range.FirstColumn;
                    dimension.LastColIndex = range.LastColumn;
                    break;
                }
            }

            bool result = false;
            if (rowIndex >= 0 && sheet.LastRowNum >= rowIndex)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row != null && columnIndex >= 0 && row.LastCellNum > columnIndex)
                {
                    ICell cell = row.GetCell(columnIndex);
                    result = cell.IsMergedCell;

                    if (dimension.Cell == null)
                    {
                        dimension.Cell = cell;
                    }
                }
            }
            dimension.IsMergeCell = result;
            return result;
        }
        #endregion
    }
}
