using System;
using System.Collections.Generic;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：导出帮助实体
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Model
{
    /// <summary>
    /// 导出帮助实体：用于传递一些参数；可根据子级需求进行扩展
    /// </summary>
    public class ExportRunEntity
    {
        private short theight = 2;//12.75
        private short cheight = 1;//25.5
        private string filename = DateTime.Now.ToString("yyyyMMddHHmmssfff");
        private ExportExcelSuffix suffix = ExportExcelSuffix.xls;
        private int sheetnum = 1;
        private List<string> sheetname = new List<string> { "Sheet1" };
        private List<string> sheettitle = new List<string>();
        private bool minusredmark = true;
        private bool autoFilter = true;
        private bool freezeTitleRow = true;
        private bool showgridline = false;
        private bool onemain = true;

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName
        {
            set { this.filename = value; }
            get { return this.filename; }
        }

        /// <summary>
        /// 标题行行高
        /// </summary>
        public short THeight
        {
            set
            {
                short dbHeight = 0;
                short.TryParse(value.ToString(), out dbHeight);
                if (dbHeight == 0)
                    dbHeight = 2;
                this.theight = (short)(dbHeight * 255);
            }
            get { return this.theight; }
        }

        /// <summary>
        /// 内容行行高
        /// </summary>
        public short CHeight
        {
            set
            {
                short dbHeight = 0;
                short.TryParse(value.ToString(), out dbHeight);
                if (dbHeight == 0)
                    dbHeight = 1;
                this.cheight = (short)(dbHeight * 255);
            }
            get { return this.cheight; }
        }

        /// <summary>
        /// Excel版本
        /// </summary>
        public ExportExcelSuffix Suffix
        {
            set { this.suffix = value; }
            get { return this.suffix; }
        }

        /// <summary>
        /// 工作表数量
        /// </summary>
        public int SheetNum
        {
            set { this.sheetnum = value; }
            get { return this.sheetnum; }
        }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public List<string> SheetName
        {
            set { this.sheetname = value; }
            get { return this.sheetname; }
        }

        public List<string> SheetTitle
        {
            set { this.sheettitle = value; }
            get { return this.sheettitle; }
        }

        public bool OneMain
        {
            set { this.onemain = value; }
            get { return this.onemain; }
        }

        /// <summary>
        /// 是否字体加粗
        /// </summary>
        public bool TitleBoldMark { set; get; }

        /// <summary>
        /// 负数标红
        /// </summary>
        public bool MinusRedMark
        {
            set { this.minusredmark = value; }
            get { return this.minusredmark; }
        }

        public bool ShowGridLine
        {
            set { this.ShowGridLine = value; }
            get { return this.showgridline; }
        }

        /// <summary>
        /// 合并列
        /// </summary>
        public int MergeColNum { set; get; }

        /// <summary>
        /// 跳过行数量
        /// </summary>
        public int SkipRowNum { set; get; }

        /// <summary>
        /// 跳过列数量
        /// </summary>
        public int SkipColNum { set; get; }

        /// <summary>
        /// 标题自动筛选过滤
        /// </summary>
        public bool AutoFilter
        {
            set { this.autoFilter = value; }
            get { return this.autoFilter; }
        }

        /// <summary>
        /// 冻结标题行
        /// </summary>
        public bool FreezeTitleRow
        {
            set { this.freezeTitleRow = value; }
            get { return this.freezeTitleRow; }
        }

        public IEnumerable<ColorEntity> ExportColors { set; get; }

        /// <summary>
        /// 用于简单列表标题行 字段、Excel类型、Excel列名、小数位
        /// </summary>
        public IEnumerable<ExportColumnEntity> ExportColumns { set; get; }

        /// <summary>
        /// 图片信息
        /// </summary>
        public IEnumerable<ExportPictureEntity> ExportPictures { set; get; }

        /// <summary>
        /// 批注信息
        /// </summary>
        public IEnumerable<CommentEntity> ExportComments { set; get; }

        /// <summary>
        /// 样式集合
        /// </summary>
        public IEnumerable<CellStyleEntity> ExportStyles { set; get; }
    }
}
