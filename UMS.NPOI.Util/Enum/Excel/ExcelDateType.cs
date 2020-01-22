
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
    /// 日期类型
    /// </summary>
    public enum ExcelDateType
    {
        /// <summary>
        /// 类型：yyyy年m月
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：2019年3月</para>
        /// </summary>
        自定义1 = 1,

        /// <summary>
        /// 类型：m月d日
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：3月23日</para>
        /// </summary>
        自定义2 = 2,

        /// <summary>
        /// 类型：yyyy年m月d日
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：2019年3月23日</para>
        /// </summary>
        自定义3 = 4,

        /// <summary>
        /// 类型：m/d/yy
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：23/3/19</para>
        /// </summary>
        自定义4 = 5,

        /// <summary>
        /// 类型：d-mmm-yy
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：23-Mar-19</para>
        /// </summary>
        自定义5 = 6,

        /// <summary>
        /// 类型：d-mmm
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：23-Mar</para>
        /// </summary>
        自定义6 = 7,

        /// <summary>
        /// 类型：mmm-yy
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：Mar-19</para>
        /// </summary>
        自定义7 = 8,

        /// <summary>
        /// 类型：h:mm AM/PM
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9.22 AM</para>
        /// </summary>
        自定义8 = 9,

        /// <summary>
        /// 类型：h:mm:ss AM/PM
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9.22:48 AM</para>
        /// </summary>
        自定义9 = 10,

        /// <summary>
        /// 类型：h:mm
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9：22</para>
        /// </summary>
        自定义10 = 11,

        /// <summary>
        /// 类型：h:mm:ss
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9:22:48</para>
        /// </summary>
        自定义11 = 12,

        /// <summary>
        /// 类型：h时mm分
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9时22分</para>
        /// </summary>
        自定义12 = 13,

        /// <summary>
        /// 类型：h时mm分ss秒
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9时22分48秒</para>
        /// </summary>
        自定义13 = 14,

        /// <summary>
        /// 类型：上午/下午h时mm分
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：上午9时22分</para>
        /// </summary>
        自定义14 = 15,

        /// <summary>
        /// 类型：上午/下午h时mm分ss秒
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：上午9时22分48秒</para>
        /// </summary>
        自定义15 = 16,

        /// <summary>
        /// 类型：yyyy/m/d h:mm
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：2019/3/23 9:22</para>
        /// </summary>
        自定义16 = 17,

        /// <summary>
        /// 类型：mm:ss
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果：9:48</para>
        /// </summary>
        自定义17 = 18,

        /// <summary>
        /// 类型：mm:ss.0
        /// <para>实际数据：2019/3/23  9:22:48</para>
        /// <para>效果:9:47.6</para>
        /// </summary>
        自定义18 = 19,

        /// <summary>
        /// 类型：yyyy/m/d
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：2019/3/24</para>
        /// </summary>
        日期1 = 20,

        /// <summary>
        /// 类型：[$-F800]dddd, mmmm dd, yyyy
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：2019年3月23日</para>
        /// </summary>
        日期2 = 21,

        /// <summary>
        /// 类型：[DBNum1][$-804]yyyy"年"m"月"d"日";@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：二〇一九年三月二十四日</para>
        /// </summary>
        日期3 = 22,

        /// <summary>
        /// 类型：[DBNum1][$-804]yyyy"年"m"月";@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：二〇一九年三月</para>
        /// </summary>
        日期4 = 23,

        /// <summary>
        /// 类型：[DBNum1][$-804]m"月"d"日";@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：三月二十四日</para>
        /// </summary>
        日期5 = 24,

        /// <summary>
        /// 类型：[$-804]aaaa;@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：星期日</para>
        /// </summary>
        日期6 = 25,

        /// <summary>
        /// 类型：[$-804]aaa;@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：周日</para>
        /// </summary>
        日期7 = 26,

        /// <summary>
        /// 类型：[DBNum1][$-804]h"时"mm"分";@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>效果：三时二十二分</para>
        /// </summary>
        时间1 = 29,

        /// <summary>
        /// 类型：[DBNum1][$-804]上午/下午h"时"mm"分";@
        /// <para>实际数据：2019/3/24  3:22:48</para>
        /// <para>上午三时二十二分</para>
        /// </summary>
        时间2 = 30,
    }
}
