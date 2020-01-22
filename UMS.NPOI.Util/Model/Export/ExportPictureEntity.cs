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
    public class ExportPictureEntity
    {
        private UrlType urltype = UrlType.Base64;
        private double scale = 1;

        public string Url { set; get; }

        public UrlType UrlType
        {
            set { this.urltype = value; }
            get { return this.urltype; }
        }

        public int RowIndex { set; get; }

        public int ColIndex { set; get; }

        public double Scale
        {
            set
            {
                double dbValue = 0;
                double.TryParse(value.ToString(), out dbValue);
                if (dbValue == 0)
                    dbValue = 1;
                this.scale = dbValue;
            }
            get { return this.scale; }
        }
    }
}
