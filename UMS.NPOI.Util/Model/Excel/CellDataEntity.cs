using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil.Model
{
    public class CellDataEntity
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public CellDataEntity()
        {

        }

        /// <summary>
        /// 标题名称
        /// </summary>
        public string TitleName { set; get; }

        /// <summary>
        /// 字段名称
        /// </summary>
        public string ColumnName { set; get; }

        /// <summary>
        /// 值
        /// </summary>
        public string CellValue { set; get; }

        /// <summary>
        /// 行位置： 0开始
        /// </summary>
        public int RowIndex { set; get; }

        /// <summary>
        /// 列位置： 0 开始
        /// </summary>
        public int ColIndex { set; get; }

        /// <summary>
        /// 单元格位置： A1
        /// </summary>
        public string CellPostion { set; get; }

        /// <summary>
        /// 字符串长度
        /// </summary>
        public int Length { set; get; }

        /// <summary>
        /// 是否为空
        /// </summary>
        public bool IsEmpty { set; get; }

        /// <summary>
        /// 转换错误
        /// </summary>
        public bool ConvertSuccess { set; get; }
    }
}
