using NPOI.XWPF.UserModel;
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
    public class TableCell
    {
        public XWPFTable Table { set; get; }

        public XWPFTableCell Cell { set; get; }

        public XWPFParagraph Paragraph { set; get; }

        public int TableIndex { set; get; }

        public int RowIndex { set; get; }

        public int CellIndex { set; get; }

        public int CellCount { set; get; }
    }
}
