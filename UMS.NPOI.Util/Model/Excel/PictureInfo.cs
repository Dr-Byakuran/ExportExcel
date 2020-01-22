using System;

/********************************************************************************
 ** 版 本：
 ** 创 建：詹建妹（James_zhan@intretech.com）
 ** 日 期：2019/03/23 16:51:55
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Model
{
    /// <summary>
    /// Excel图片信息实体
    /// </summary>
    public class PictureInfo
    {
        /// <summary>
        /// 开始行位置
        /// </summary>
        public int FirstRow { get; set; }

        /// <summary>
        /// 结束行位置
        /// </summary>
        public int LastRow { get; set; }

        /// <summary>
        /// 开始列位置
        /// </summary>
        public int FirstCol { get; set; }

        /// <summary>
        /// 结束列位置
        /// </summary>
        public int LastCol { get; set; }

        /// <summary>
        /// 图片后缀
        /// </summary>
        public string Suffix { set; get; }

        /// <summary>
        /// 图片数据
        /// </summary>
        public Byte[] PictureData { get; private set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="firstRow">开始行</param>
        /// <param name="lastRow">结束行</param>
        /// <param name="firstCol">开始列</param>
        /// <param name="lastCol">结束列</param>
        /// <param name="suffix">图片后缀</param>
        /// <param name="pictureData">图片数据</param>
        public PictureInfo(int firstRow, int lastRow, int firstCol, int lastCol, string suffix, Byte[] pictureData)
        {
            this.FirstRow = firstRow;
            this.LastRow = lastRow;
            this.FirstCol = firstCol;
            this.LastCol = lastCol;
            this.PictureData = pictureData;
            this.Suffix = suffix;
        }
    }
}
