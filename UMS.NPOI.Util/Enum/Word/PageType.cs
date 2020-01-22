
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
    /// 页面类型
    /// </summary>
    public enum PageType
    {
        /// <summary>
        /// 信纸
        /// </summary>
        LetterPager_W = 1,

        /// <summary>
        /// 信纸
        /// </summary>
        LetterPager_H = 2,

        /// <summary>
        /// A4
        /// </summary>
        A4_W = 3,

        /// <summary>
        /// A4
        /// </summary>
        A4_H = 4,

        /// <summary>
        /// 16 开
        /// </summary>
        K16_W = 5,

        /// <summary>
        /// 16 开
        /// </summary>
        K16_H = 6,

        /// <summary>
        /// 32 开 
        /// </summary>
        K32_W = 7,

        /// <summary>
        /// 32 开
        /// </summary>
        K32_H = 8,

        /// <summary>
        /// 大32开
        /// </summary>
        KB32_W = 9,

        /// <summary>
        /// 大32开
        /// </summary>
        KB32_H = 10,

        /// <summary>
        /// 自定义纸张大小
        /// </summary>
        Custom = 6,
    }
}
