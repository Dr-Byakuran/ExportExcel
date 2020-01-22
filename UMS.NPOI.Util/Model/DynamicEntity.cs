using System;
using System.Collections.Generic;
using System.Dynamic;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：动态实体
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Model
{
    /// <summary>
    /// 动态实体
    /// </summary>
    public class DynamicEntity : DynamicObject
    {
        private Dictionary<string, object> _value;

        /// <summary>
        /// 构造函数
        /// </summary>
        public DynamicEntity()
        {
            _value = new Dictionary<string, object>();
        }

        /// <summary>
        /// 值集合
        /// </summary>
        public Dictionary<string, object> Value
        {
            set { this.Value = _value; }
            get { return this._value; }
        }

        /// <summary>
        /// 取值/赋值
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public object this[string key]
        {
            set
            {
                SetPropertyValue(key, value);
            }
            get
            {
                return GetPropertyValue(key);
            }
        }

        /// <summary>
        /// 获取属性值
        /// </summary>
        /// <param name="propertyName">属性名称</param>
        /// <param name="ordinalIgnoreCase">是否忽略大小写：默认忽略</param>
        /// <returns></returns>
        public object GetPropertyValue(string propertyName, bool ignoreCase = true)
        {
            string name = propertyName;
            if (ignoreCase)
                name = name.ToLower();
            if (_value.ContainsKey(name) == true)
                return _value[name];
            else
                return null;
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        /// <param name="propertyName"></param>
        /// <param name="value"></param>
        public void SetPropertyValue(string propertyName, object obj)
        {
            if (_value.ContainsKey(propertyName) == true)
            {
                _value[propertyName] = obj;
            }
            else
            {
                _value.Add(propertyName, obj);
            }
        }

        public bool IsEntityProperty(string propertyName)
        {
            bool rs = true;
            rs = _value.ContainsKey(propertyName);
            return rs;
        }

        /// <summary>
        /// 实现动态对象属性成员访问的方法，得到返回指定属性的值
        /// </summary>
        /// <param name="binder"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = GetPropertyValue(binder.Name);
            return result == null ? false : true;
        }

        /// <summary>
        /// 实现动态对象属性值设置的方法。
        /// </summary>
        /// <param name="binder"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            SetPropertyValue(binder.Name, value);
            return true;
            //return base.TrySetMember(binder, value);
        }

        /// <summary>
        /// 动态对象动态方法调用时执行的实际代码
        /// </summary>
        /// <param name="binder"></param>
        /// <param name="args"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            var theDelegateObj = GetPropertyValue(binder.Name) as DelegateObj;
            if (theDelegateObj == null || theDelegateObj.CallMethod == null)
            {
                result = null;
                return false;
            }
            result = theDelegateObj.CallMethod(this, args);
            return true;
            //return base.TryInvokeMember(binder, args, out result);
        }

        public override bool TryInvoke(InvokeBinder binder, object[] args, out object result)
        {
            return base.TryInvoke(binder, args, out result);
        }
    }
    /// <summary>
    /// 自定义委托
    /// </summary>
    /// <param name="Sender"></param>
    /// <param name="PMs"></param>
    /// <returns></returns>
    public delegate object MyDelegate(dynamic Sender, params object[] PMs);

    /// <summary>
    /// 委托实体
    /// </summary>
    public class DelegateObj
    {
        private MyDelegate _delegate;

        public MyDelegate CallMethod
        {
            get { return _delegate; }
        }
        private DelegateObj(MyDelegate D)
        {
            _delegate = D;
        }
        /// <summary>
        /// 构造委托对象，让它看起来有点javascript定义的味道.
        /// </summary>
        /// <param name="D"></param>
        /// <returns></returns>
        public static DelegateObj Function(MyDelegate D)
        {
            return new DelegateObj(D);
        }
    }

}
