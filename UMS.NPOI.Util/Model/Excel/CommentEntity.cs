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
    public class CommentEntity
    {
        private int dx1 = 0;
        private int dx2 = 0;
        private int dy1 = 0;
        private int dy2 = 0;
        private int anchortype = 2;

        /// <summary>
        /// 批注内容
        /// </summary>
        public string Text { set; get; }

        /// <summary>
        /// 行位置
        /// </summary>
        public int RowIndex { set; get; }

        /// <summary>
        /// 列位置
        /// </summary>
        public int ColIndex { set; get; }

        /// <summary>
        /// 批注 所占行
        /// </summary>
        public int Height { set; get; }

        /// <summary>
        /// 批注 所占列
        /// </summary>
        public int Width { set; get; }

        /// <summary>
        /// 锚 类型
        /// </summary>
        public int AnchorType
        {
            set { this.anchortype = value; }
            get { return this.anchortype; }
        }

        /// <summary>
        /// 开始列 偏移量
        /// </summary>
        public int Dx1
        {
            set { this.dx1 = value; }
            get { return this.dx1; }
        }

        /// <summary>
        /// 结束列 偏移量
        /// </summary>
        public int Dx2
        {
            set { this.dx2 = value; }
            get { return this.dx2; }
        }

        /// <summary>
        /// 开始行 偏移量
        /// </summary>
        public int Dy1
        {
            set { this.dy1 = value; }
            get { return this.dy1; }
        }

        /// <summary>
        /// 结束行 偏移量
        /// </summary>
        public int Dy2
        {
            set { this.dy2 = value; }
            get { return this.dy2; }
        }
       
    }
}
