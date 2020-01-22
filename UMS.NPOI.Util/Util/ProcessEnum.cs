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
namespace UMS.Framework.NpoiUtil.Util
{
    public static class ProcessEnum
    {
        public static int ToInt(this Type type)
        {
            int intType = -1;
            intType = (int)Enum.Parse(typeof(Type), type.ToString());
            
            return intType;
        }
    }
}
