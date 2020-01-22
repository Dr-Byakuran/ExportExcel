using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil.Model
{
    public class CellImportEntity
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public CellImportEntity()
        {

        }
        /// <summary>
        /// 数据字段名称
        /// </summary>
        public string ColumnName { set; get; }

        /// <summary>
        /// 数据所在行位置
        /// </summary>
        public int RowIndex { set; get; } 

        /// <summary>
        /// 数据坐在列位置
        /// </summary>
        public int CellIndex { set; get; }

        /// <summary>
        /// 是否何必列
        /// </summary>
        public bool IsMerge { set; get; }

        /// <summary>
        /// 合并开始行位置
        /// </summary>
        public int? MergeBeginRow { set; get; }

        /// <summary>
        /// 合并开始列位置
        /// </summary>
        public int? MergeBeginCell { set; get; }

        /// <summary>
        /// 合并结束行位置
        /// </summary>
        public int? MergeEndRow { set; get; }

        /// <summary>
        /// 合并结束列位置
        /// </summary>
        public int? MergeEndCell { set; get; }
    }
}
