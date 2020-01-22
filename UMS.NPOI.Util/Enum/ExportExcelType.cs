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
    /// <summary>
    /// 导出Excel 类型
    /// </summary>
    public enum ExportExcelType
    {
        /// <summary>
        /// 简单列表导出
        /// </summary>
        Simple = 1,

        /// <summary>
        /// 对账类型
        /// </summary>
        Reconciliation = 2,

        /// <summary>
        /// 标题合并类型
        /// </summary>
        Merge = 3,

        /// <summary>
        /// 单据
        /// </summary>
        Bill = 4,
    }
}
