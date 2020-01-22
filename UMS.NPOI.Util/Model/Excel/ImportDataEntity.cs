using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UMS.Framework.NpoiUtil.Model
{
    public  class ImportDataEntity
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ImportDataEntity()
        {

        }

        /// <summary>
        /// 导入的数据：横向数据
        /// </summary>
        public object ParemtEntity { set; get; }

        /// <summary>
        /// 单元格数据信息，用于横向数据验证
        /// </summary>
        public IEnumerable<CellDataEntity> ParentCells { set; get; }

        /// <summary>
        /// 导入的数据集合：纵向数据
        /// </summary>
        public IEnumerable<object> ChildEntity { set; get; }

        /// <summary>
        /// 单元格数据信息，用于纵向数据验证
        /// </summary>
        public IEnumerable<CellDataEntity> ChildCells { set; get; }
    }
}
