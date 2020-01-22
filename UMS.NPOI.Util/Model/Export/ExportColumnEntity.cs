using NPOI.SS.UserModel;
using System.Collections.Generic;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：导出Excel字段帮助实体
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Model
{
    /// <summary>
    /// 导出Excel字段帮助实体
    /// </summary>
    public class ExportColumnEntity
    {
        private CellType cellType = CellType.String;
        private HorizontalAlignment halign = HorizontalAlignment.Left;
        private VerticalAlignment valign = VerticalAlignment.Center;
        private int dot = 2;
        private double width = 8.43;
        private MergeAlign mergealgin = MergeAlign.LR;
        private int mergerow = 0;
        private int mergecol = 0;

        #region 简单操作参数

        /// <summary>
        /// 字段名
        /// </summary>
        public string ColumnName { set; get; }

        /// <summary>
        /// Excel显示名称
        /// </summary>
        public string ExcelName { set; get; }

        /// <summary>
        /// 当前数据类型
        /// </summary>
        public CellType CellType
        {
            set { this.cellType = value; }
            get { return this.cellType; }
        }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public HorizontalAlignment HAlign
        {
            set { this.halign = value; }
            get { return this.halign; }
        }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public VerticalAlignment VAlign
        {
            set { this.valign = value; }
            get { return this.valign; }
        }

        /// <summary>
        /// 宽度
        /// </summary>
        public int Width
        {
            set
            {
                double dbWidth = 0;
                double.TryParse(value.ToString(), out dbWidth);
                if (dbWidth == 0)
                    dbWidth= 8.43;
                this.width = ((dbWidth + 0.72) * 256);
            }
            get { return (int)this.width; }
        }

        /// <summary>
        /// 自定义格式
        /// </summary>
        public string DataFormat { set; get; }

        /// <summary>
        /// 小数位数：配合数值类型使用
        /// </summary>
        public int Dot
        {
            set { this.dot = value; }
            get { return this.dot; }
        }

        /// <summary>
        /// 字符串转换
        /// </summary>
        public Dictionary<string, string> Dic { set; get; }

        /// <summary>
        /// 是否隐藏，不导出
        /// </summary>
        public bool Hidden { set; get; }

        /// <summary>
        /// 字典分类编码
        /// </summary>
        public string DataItemCode { set; get; }

        #endregion

        #region 复杂操作参数 
        /// <summary>
        /// 别名
        /// </summary>
        public string Alias { set; get; }

        /// <summary>
        /// 用于区别主从字段
        /// </summary>
        public bool PrimaryMark { set; get; }

        /// <summary>
        /// 行位置
        /// </summary>
        public int RowIndex { set; get; }

        /// <summary>
        /// 列位置
        /// </summary>
        public int ColIndex { set; get; }

        /// <summary>
        /// 合并后标题名称
        /// </summary>
        public string MergeName { set; get; }

        /// <summary>
        /// 批注标识
        /// </summary>
        public bool NoteMark { set; get; }

        /// <summary>
        /// 批注内容
        /// </summary>
        public string NoteContent { set; get; }

        /// <summary>
        /// 标题列值所占行
        /// </summary>
        public int TitleColSpan { set; get; }

        /// <summary>
        /// 差异数
        /// </summary>
        public int diffNum { set; get; }

        /// <summary>
        /// 标题合并方向
        /// </summary>
        public MergeAlign MergeAlign
        {
            set { this.mergealgin = value; }
            get { return this.mergealgin; }
        }

        /// <summary>
        /// 合并行
        /// </summary>
        public int MergeRow
        {
            set { this.mergerow = value; }
            get { return this.mergerow; }
        }

        /// <summary>
        /// 合并列
        /// </summary>
        public int MergeCol
        {
            set { this.mergecol = value; }
            get { return this.mergecol; }
        }

        #endregion

        /// <summary>
        /// 样式标识：用于获取样式
        /// </summary>
        public int CellStyleIndex { set; get; }

        /// <summary>
        /// 根据样式标识，获取的样式
        /// </summary>
        public ICellStyle CellStyle { set; get; }
    }
}
