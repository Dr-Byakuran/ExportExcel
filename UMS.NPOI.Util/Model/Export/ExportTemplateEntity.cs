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
    public class ExportTemplateEntity
    {
        private PathType pathtype = PathType.Http;
        /// <summary>
        /// 文件名称
        /// </summary>
        public string FileName { set; get; }

        /// <summary>
        /// 地址
        /// </summary>
        public string Path { set; get; }

        /// <summary>
        /// 地址类型
        /// </summary>
        public PathType PathType
        {
            set { this.pathtype = value; }
            get { return this.pathtype; }
        }

        /// <summary>
        /// 后缀名类型
        /// </summary>
        public ExportExcelSuffix Suffix { set; get; }

        /// <summary>
        /// 模板类型
        /// </summary>
        public TemplateType TemplateType { set; get; }
    }
}
