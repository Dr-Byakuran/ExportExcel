using NPOI.SS.UserModel;

/********************************************************************************
 ** 版 本：
 ** 创 建：詹建妹（James_zhan@intretech.com）
 ** 日 期：2019/03/23 16:51:55
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Model
{
    /// <summary>
    /// 重定义名称实体
    /// </summary>
    public class INameInfo
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { set; get; }

        /// <summary>
        /// 工作表位置
        /// </summary>
        public int SheetIndex { set; get; }

        /// <summary>
        /// 说明备注
        /// </summary>
        public string Comment { set; get; }

        /// <summary>
        /// 是否已被删除
        /// </summary>
        public bool IsDeleted { set; get; }

        /// <summary>
        /// 是否是函数名
        /// </summary>
        public bool IsFuntionName { set; get; }

        /// <summary>
        /// 重定义名称
        /// </summary>
        public string NameName { set; get; }

        /// <summary>
        /// 索引地址
        /// </summary>
        public string RefersToFormula { set; get; }

        /// <summary>
        /// 开始行
        /// </summary>
        public int FirstRow { set; get; }

        /// <summary>
        /// 结束行
        /// </summary>
        public int LastRow { set; get; }

        /// <summary>
        /// 开始列
        /// </summary>
        public int FirstCol { set; get; }

        /// <summary>
        /// 结束行
        /// </summary>
        public int LastCol { set; get; }

        /// <summary>
        /// 行一致
        /// </summary>
        public bool EqualRow { set; get; }

        /// <summary>
        /// 列一致
        /// </summary>
        public bool EqualCol { set; get; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="iname"></param>
        public INameInfo(IName iname)
        {
            this.SheetName = iname.SheetName;
            this.SheetIndex = iname.SheetIndex;
            this.Comment = iname.Comment;
            this.IsDeleted = iname.IsDeleted;
            this.IsFuntionName = iname.IsFunctionName;
            this.NameName = iname.NameName;
            this.RefersToFormula = iname.RefersToFormula;
        }
    }
}
