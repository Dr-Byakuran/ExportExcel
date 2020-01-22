
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
    /// Excel 边框线类型
    /// </summary>
    public enum ExcelBorderType
    {
        /// <summary>
        /// 下框线
        /// </summary>
        BorderBottom = 1,

        /// <summary>
        /// 上框线
        /// </summary>
        BorderTop = 2,

        /// <summary>
        /// 左框线
        /// </summary>
        BorderLeft = 3,

        /// <summary>
        /// 右框线
        /// </summary>
        BorderRight = 4,

        /// <summary>
        /// 所有框线
        /// </summary>
        BorderAll = 5,

        /// <summary>
        /// 所有框线(加粗)
        /// </summary>
        BorderAllBold = 6,

        /// <summary>
        /// 双底框线
        /// </summary>
        BorderBottomDouble = 7,

        /// <summary>
        /// 粗底框线
        /// </summary>
        BorderBottomBold = 8,

        /// <summary>
        /// 上下框线
        /// </summary>
        BorderTopAndBottom = 9,

        /// <summary>
        /// 上框线和下粗框线
        /// </summary>
        BorderTopAndBotoomBold = 10,

        /// <summary>
        /// 上框线和双下框线
        /// </summary>
        BorderTopAndBottomDouble = 11,
    }
}
