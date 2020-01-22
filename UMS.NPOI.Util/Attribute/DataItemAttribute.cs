using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// 数据字典标识扩展属性
    /// </summary>
    public class DataItemAttribute: Attribute
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public DataItemAttribute()
        {

        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="itemcode"></param>
        /// <param name="itemvalue"></param>
        public DataItemAttribute(string itemcode, bool itemvalue = true)
        {
            this.ItemCode = itemcode;
            this.ItemValue = itemvalue;
        }

        public string ItemCode { set; get; }

        /// <summary>
        /// 是否 根据名称取值：默认是，否则根据值取名称
        /// </summary>
        public bool ItemValue { set; get; }
    }
}
