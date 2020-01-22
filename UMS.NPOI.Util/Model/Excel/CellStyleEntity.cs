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
namespace UMS.Framework.NpoiUtil.Model
{
    public class CellStyleEntity
    {
        private HorizontalAlignment alignment = HorizontalAlignment.Left;
        private VerticalAlignment verticalAlignment = VerticalAlignment.Center;
        private BorderStyle borderTop = BorderStyle.Thin;
        private BorderStyle borderRight = BorderStyle.Thin;
        private BorderStyle borderBottom = BorderStyle.Thin;
        private BorderStyle borderLeft = BorderStyle.Thin;
        private BorderDiagonal borderDiagonal = BorderDiagonal.None;
        private bool wrapText = true;

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public HorizontalAlignment Alignment
        {
            set { this.alignment = value; }
            get { return this.alignment; }
        }

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            set { this.verticalAlignment = value; }
            get { return this.verticalAlignment; }
        }

        /// <summary>
        /// 上边框 样式
        /// </summary>
        public BorderStyle BorderTop
        {
            set { this.borderTop = value; }
            get { return this.borderTop; }
        }

        /// <summary>
        /// 有边框 样式
        /// </summary>
        public BorderStyle BorderRight
        {
            set { this.borderRight = value; }
            get { return this.borderRight; }
        }

        /// <summary>
        /// 下边框 样式
        /// </summary>
        public BorderStyle BorderBottom
        {
            set { this.borderBottom = value; }
            get { return this.borderBottom; }
        }

        /// <summary>
        /// 左边框 样式
        /// </summary>
        public BorderStyle BorderLeft
        {
            set { this.borderLeft = value; }
            get { return this.borderLeft; }
        }

        /// <summary>
        /// 边框对角线
        /// </summary>
        public BorderDiagonal BorderDiagonal
        {
            set { this.borderDiagonal = value; }
            get { return this.borderDiagonal; }
        }

        /// <summary>
        /// 边框对角线 颜色
        /// </summary>
        public string BorderDiagonalColor
        {
            set;get;
        }

        /// <summary>
        /// 顶部边框  颜色
        /// </summary>
        public string TopBorderColor
        {
            set; get;
        }

        /// <summary>
        /// 右边边框 颜色
        /// </summary>
        public string RightBorderColor
        {
            set; get;
        }

        /// <summary>
        /// 左边边框 颜色
        /// </summary>
        public string LeftBorderColor
        {
            set; get;
        }

        /// <summary>
        /// 底部边框 颜色
        /// </summary>
        public string BottomBorderColor
        {
            set; get;
        }

        /// <summary>
        /// 填充背景色
        /// </summary>
        public string FillBackgroundColor
        {
            set; get;
        }

        /// <summary>
        /// 填充前景色
        /// </summary>
        public string FillForegroundColor
        {
            set; get;
        }

        /// <summary>
        /// 边框对角线 样式
        /// </summary>
        public BorderStyle BorderDiagonalLineStyle
        {
            set;get;
        }

        /// <summary>
        /// 格式化
        /// </summary>
        public string DataFormat
        {
            set;get;
        }

        /// <summary>
        /// 填充模式
        /// </summary>
        public FillPattern FillPattern
        {
            set;get;
        }

        /// <summary>
        /// 缩进
        /// </summary>
        public short Indention
        {
            set;get;
        }

        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool IsHidden
        {
            set;get;
        }

        /// <summary>
        /// 是否锁定
        /// </summary>
        public bool IsLocked
        {
            set;get;
        }

        /// <summary>
        /// 是否旋转
        /// </summary>
        public short Rotation
        {
            set;get;
        }

        /// <summary>
        /// 是否缩进适应
        /// </summary>
        public bool ShrinkToFit
        {
            set;get;
        }

        /// <summary>
        /// 自动换行
        /// </summary>
        public bool WrapText
        {
            set { this.wrapText = value; }
            get { return this.wrapText; }
        }

        /// <summary>
        /// 字体
        /// </summary>
        public FontEntity Font { set; get; }

        /// <summary>
        /// 样式标识
        /// </summary>
        public int CellStyleIndex { set; get; }

        /// <summary>
        /// 根据样式标识 获取样式
        /// </summary>
        public ICellStyle CellStyle { set; get; }
    }
}
