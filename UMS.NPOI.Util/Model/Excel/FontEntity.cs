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
    public class FontEntity
    {
        private FontUnderlineType underline = FontUnderlineType.None;

        /// <summary>
        /// 加粗权重
        /// </summary>
        public short Boldweight
        {
            set;get;
        }

        /// <summary>
        /// 字符集
        /// </summary>
        public short Charset
        {
            set;get;
        }

        /// <summary>
        /// 颜色
        /// </summary>
        public string Color
        {
            set;get;
        }

        /// <summary>
        /// 字体高
        /// </summary>
        public double FontHeight
        {
            set;get;
        }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short FontHeightInPoints
        {
            set;get;
        }

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName
        {
            set;get;
        }

        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool IsItalic
        {
            set;get;
        }

        /// <summary>
        /// 是否添加删除线
        /// </summary>
        public bool IsStrikeout
        {
            set;get;
        }

        /// <summary>
        /// 字体上标
        /// </summary>
        public FontSuperScript TypeOffset
        {
            set;get;
        }

        /// <summary>
        /// 字体下 线类型
        /// </summary>
        public FontUnderlineType Underline
        {
            set { this.underline = value; }
            get { return this.underline; }
        }
    }
}
