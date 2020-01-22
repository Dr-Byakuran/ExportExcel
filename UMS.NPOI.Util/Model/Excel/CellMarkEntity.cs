using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil.Model
{
    public class CellMarkEntity
    {
        private bool transgain = false;
        /// <summary>
        /// 构造函数
        /// </summary>
        public CellMarkEntity()
        {

        }

        /// <summary>
        /// 标题名称
        /// </summary>
        public string TitleName { set; get; }

        /// <summary>
        /// 标记字段 名称
        /// </summary>
        public string ColumnName { set; get; }

        /// <summary>
        /// 标记字段 行位置
        /// </summary>
        public int RowIndex { set; get; }

        /// <summary>
        /// 标记字段 列位置
        /// </summary>
        public int CellIndex { set; get; }

        /// <summary>
        /// 标记字段 IName 属性
        /// </summary>
        public IName IName { set; get; }

        /// <summary>
        /// 是否横向获取数据
        /// </summary>
        public bool TransGain
        {
            set { this.transgain = value; }
            get { return this.transgain; }
        }

        /// <summary>
        /// 属性
        /// </summary>
        public PropertyInfo PropertyInfo { set; get; }

        /// <summary>
        /// 自定义属性
        /// </summary>
        public Attribute Attribute { set; get; }
    }
}
