using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using UMS.Framework.NpoiUtil.Model;
using System.Data;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Util
{
    public static class DynamicUtil
    {
        /// <summary>
        /// 实体集合 转 动态实体集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static IEnumerable<DynamicEntity> ProcessModel<T>(this IEnumerable<T> list)
            where T : class, new()
        {
            List<DynamicEntity> newList = new List<DynamicEntity>();
            DynamicEntity newEntity = null ;
            foreach(var entity in list)
            {
                newEntity = entity.ProcessModel();
                newList.Add(newEntity);
            }
            return newList.AsEnumerable();
        }

        /// <summary>
        /// 实体 转 动态实体
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <returns></returns>
        public static DynamicEntity ProcessModel<T>(this T entity)
            where T : class, new()
        {
            DynamicEntity newEntity = new DynamicEntity();
            PropertyInfo[] infos = entity.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var info in infos)
            {
                string propertyName = info.Name;
                object value = entity.GetType().GetProperty(propertyName).GetValue(entity, null);
                newEntity.SetPropertyValue(propertyName, value);
            }
            return newEntity;
        }

        /// <summary>
        /// DataTable 转 动态实体集合
        /// </summary>
        /// <param name="dtData"></param>
        /// <returns></returns>
        public static IEnumerable<DynamicEntity> ProcessTable(this DataTable dtData)
        {
            List<DynamicEntity> list = new List<DynamicEntity>();
            if (dtData == null || dtData.Rows.Count == 0)
                return list;
            DynamicEntity entity = null;
            foreach (DataRow dr in dtData.Rows)
            {
                entity = new DynamicEntity();
                foreach (DataColumn item in dtData.Columns)
                {
                    string propertyName = item.ColumnName;
                    object value = dr[propertyName];
                    entity.SetPropertyValue(propertyName, value);
                }
                list.Add(entity);
            }
            return list;
        }
    }
}
