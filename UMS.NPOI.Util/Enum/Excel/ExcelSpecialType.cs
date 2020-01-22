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
namespace UMS.Framework.NpoiUtil
{
    public enum ExcelSpecialType
    {
        /// <summary>
        /// 类型：000000
        /// <para>邮政编码</para>
        /// </summary>
        特殊1 = 1,

        /// <summary>
        /// 类型：[DBNum1][$-804]G/通用格式
        /// <para>中文小写数字</para>
        /// <para>效果：100->一百</para>
        /// </summary>
        特殊2 = 2,

        /// <summary>
        /// 类型：[DBNum2][$-804]G/通用格式
        /// <para>中文大写数字</para>
        /// <para>效果：100->壹佰</para>
        /// </summary>
        特殊3 = 3,
    }
}
