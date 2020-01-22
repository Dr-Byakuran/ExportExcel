using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil.Model
{
    public class CellMergeEntity
    {
        public int FirstColumn { get; set; }

        public int FirstRow { get; set; }

        public int LastColumn { get; set; }

        public int LastRow { get; set; }

        public ICell DataCell { get; set; }
    }
}
