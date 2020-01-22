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
    public class PictureEntity
    {
        /// <summary>
        /// 宽度缩放比例
        /// </summary>
        public double ScaleX { set; get; }

        /// <summary>
        /// 高度缩放比例
        /// </summary>
        public double ScaleY { set; get; }

        /// <summary>
        /// 锚
        /// </summary>
        public IClientAnchor Anchor { set; get; }

        /// <summary>
        /// 图片位置
        /// </summary>
        public int PictureIndex { set; get; }

        /// <summary>
        /// 是否原尺寸，无需缩放
        /// </summary>
        public bool OriginalSize { set; get; }
    }
}
