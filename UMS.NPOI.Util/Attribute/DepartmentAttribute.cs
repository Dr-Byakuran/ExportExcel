using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// 部门标识扩展属性
    /// </summary>
    public class DepartmentAttribute : Attribute
    {
        private bool curaccountset = true;
        /// <summary>
        /// 构造函数
        /// </summary>
        public DepartmentAttribute()
        {

        }

        /// <summary>
        /// 是否获取当前账套下部门数据：默认是
        /// </summary>
        public bool CurAccountSet
        {
            set { this.curaccountset = value; }
            get { return this.curaccountset; }
        }

        /// <summary>
        /// 部门名称赋值字段
        /// </summary>
        public string DepartNameColumn { set; get; }

        /// <summary>
        /// 部门主键赋值字段
        /// </summary>
        public string DepartIdColumn { set; get; }
    }
}
