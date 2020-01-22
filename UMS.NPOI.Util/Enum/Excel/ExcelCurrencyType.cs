
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
    /// 货币类型
    /// </summary>
    public enum ExcelCurrencyType
    {
        无 = 1,

        /// <summary>
        /// ￥
        /// </summary>
        人民币 = 2,

        /// <summary>
        /// $
        /// </summary>
        美元 = 3,

        /// <summary>
        /// €
        /// </summary>
        欧元 = 4,

        /// <summary>
        /// £
        /// </summary>
        英镑 = 5,
    }
}
