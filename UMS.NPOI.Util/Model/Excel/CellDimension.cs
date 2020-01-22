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
    public class CellDimension
    {
        /// <summary>
        /// 含有数据的单元格(通常表示合并单元格的第一个跨度行第一个跨度列)，该字段可能为null
        /// </summary>
        public ICell Cell;

        /// <summary>
        /// 行跨度(跨越了多少行)
        /// </summary>
        public int RowSpan;

        /// <summary>
        /// 列跨度(跨越了多少列)
        /// </summary>
        public int ColSpan;

        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int FirstRowIndex;

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int LastRowIndex;

        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int FirstColIndex;

        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int LastColIndex;

        /// <summary>
        /// 是否合并列
        /// </summary>
        public bool IsMergeCell;
    }
}
