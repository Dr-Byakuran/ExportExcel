using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// 布尔型标识属性
    /// </summary>
    public class BooleanAttribute : Attribute
    {
        private string truevalue = "Yes";
        /// <summary>
        /// 构造函数
        /// </summary>
        public BooleanAttribute()
        {

        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="trueValue"></param>
        public BooleanAttribute(string trueValue)
            :base()
        {
            this.truevalue = trueValue;
        }

        /// <summary>
        /// 默认值 Yes
        /// </summary>
        public string TrueValue
        {
            set { this.truevalue = value; }
            get { return this.truevalue; }
        }
    }
}
