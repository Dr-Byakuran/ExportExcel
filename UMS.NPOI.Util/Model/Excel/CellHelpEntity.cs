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
    public class CellHelpEntity : CellDimension
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="entity"></param>
        public CellHelpEntity(CellDimension entity)
        {
            Cell = entity.Cell;
            RowSpan = entity.RowSpan;
            ColSpan = entity.ColSpan;
            FirstColIndex = entity.FirstColIndex;
            FirstRowIndex = entity.FirstRowIndex;
            LastColIndex = entity.LastColIndex;
            LastRowIndex = entity.LastRowIndex;
            IsMergeCell = entity.IsMergeCell;
        }

        /// <summary>
        /// 字段名称
        /// </summary>
        public string Name { set; get; }

        /// <summary>
        /// 类型
        /// </summary>
        public ConfigType Type { set; get; }

        /// <summary>
        /// 原行
        /// </summary>
        public IRow SorceRow { set; get; }
    }
}
