using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil.Util
{
    /// <summary>
    /// EXCEL 帮助类
    /// </summary>
    public class ExportExcelUtil
    {
        /// <summary>
        /// 数据集合转 DataSet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataSet ToDataSet<T>(IEnumerable<T> list)
        {
            Type elementType = typeof(T);
            var ds = new DataSet();
            var t = new DataTable();
            ds.Tables.Add(t);
            elementType.GetProperties().ToList().ForEach(propInfo => t.Columns.Add(propInfo.Name, Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType));
            foreach (T item in list)
            {
                var row = t.NewRow();
                elementType.GetProperties().ToList().ForEach(propInfo => row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value);
                t.Rows.Add(row);
            }
            return ds;
        }
        /// <summary>
        /// 获取字体
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public static short FontSize(string size)
        {
            short fontSize = 0;
            switch (size)
            {
                case "小五":
                    fontSize = 9;
                    break;
                case "五号":
                    fontSize = (short)10.5;
                    break;
                case "小四":
                    fontSize = 12;
                    break;
                case "四号":
                    fontSize = 14;
                    break;
                case "小三":
                    fontSize = 15;
                    break;
                case "三号":
                    fontSize = 16;
                    break;
                case "小二":
                    fontSize = 18;
                    break;
                case "二号":
                    fontSize = 22;
                    break;
                case "小一":
                    fontSize = 24;
                    break;
                case "一号":
                    fontSize = 26;
                    break;
                case "小初":
                    fontSize = 36;
                    break;
                case "初号":
                    fontSize = 42;
                    break;
                default:
                    fontSize = 12;
                    break;
            }
            return fontSize;
        }

        /// <summary>
        /// 对齐方式
        /// </summary>
        /// <param name="align"></param>
        /// <returns></returns>
        public static HorizontalAlignment FontAlign(string align)
        {
            HorizontalAlignment hAlign = new HorizontalAlignment();
            switch (align.ToLower())
            {
                case "left":
                    hAlign = HorizontalAlignment.Left;
                    break;
                case "center":
                    hAlign = HorizontalAlignment.Center;
                    break;
                case "right":
                    hAlign = HorizontalAlignment.Right;
                    break;
            }
            return hAlign;
        }

        /// <summary>
        /// 数值转换
        /// </summary>
        /// <returns></returns>
        public static string TransValue(string value, string formater, string option)
        {
            string result = string.Empty;
            decimal dcmTemp = 0;
            DateTime dt = new DateTime();
            if (string.IsNullOrEmpty(value))
                return result;
            switch (formater)
            {
                case "":
                    result = value;
                    break;
                case "decimal":
                    int intDecial = 0;
                    int.TryParse(option, out intDecial);
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("{0:N" + intDecial + "}", dcmTemp);
                    break;
                case "num":
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("{0:N2}", dcmTemp);
                    break;
                case "$num":
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("{0:c2}", dcmTemp);
                    break;
                case "￥num":
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("￥{0:N2}", dcmTemp);
                    break;
                case "€num":
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("€{0:N2}", dcmTemp);
                    break;
                case "num%":
                    decimal.TryParse(value, out dcmTemp);
                    result = string.Format("{0:P}", dcmTemp);
                    break;
                case "date":
                    dt = DateTime.Parse(value);
                    result = dt.ToString("yyyy-MM-dd");
                    break;
                case "dateTime":
                    dt = DateTime.Parse(value);
                    result = dt.ToString("yyyy-MM-dd hh:mm:ss");
                    break;
                case "bool":
                    if (value == "1" || value == "True")
                        result = "是";
                    else
                        result = "否";
                    break;
                case "lower":
                    result = value.ToLower();
                    break;
                case "upper":
                    result = value.ToUpper();
                    break;
                case "fixedValue":
                    result = option;
                    break;
                case "select":
                    string[] list = option.Split(';');
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    foreach (var item in list)
                    {
                        string[] obj = item.Split(':');
                        if (obj.Length == 2)
                            dic.Add(obj[0], obj[1]);
                    }
                    if (dic.Keys.Contains(value))
                    {
                        result = dic[value];
                    }
                    else
                        result = value;
                    break;
            }

            return result;
        }

        /// <summary>
        /// 将Sheet列号变为列名
        /// @param index 列号, 从0开始
        /// @return 0->A; 1->B...26->AA
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static string IndexToColName(int index)
        {
            if (index < 0)
            {
                return null;
            }
            int num = 65;// A的Unicode码
            string colName = "";
            do
            {
                if (colName.Length > 0)
                {
                    index--;
                }
                int remainder = index % 26;
                colName = ((char)(remainder + num)) + colName;
                index = (int)((index - remainder) / 26);
            } while (index > 0);
            return colName;
        }

        /// <summary>
        /// 根据表元的列名转换为列号
        /// </summary>
        /// <param name="colName"></param>
        /// <returns></returns>
        public static int ColNameToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index - 1;
        }

        /// <summary>
        /// 转换数据类型：自定义查询 DataTable 类型返回数据类型 
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        private static string GetDataType(string dataType)
        {
            var result = "";
            switch (dataType)
            {
                case "DateTime":
                    result = "datetime";
                    break;
                default:
                    result = "varchar";
                    break;
            }
            return result;
        }

        /// <summary>
        /// 获取数据长度
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int GetStrLen(string value)
        {
            byte[] bytes = System.Text.Encoding.Default.GetBytes(value);
            return bytes.Length;
        }

        public static byte[] GetHttpFile(string url)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
            byte[] bytes = new byte[webResponse.ContentLength];
            if (webResponse.StatusCode == HttpStatusCode.OK)
            {
                System.IO.Stream st = webResponse.GetResponseStream();
                st.Read(bytes, 0, Convert.ToInt32(webResponse.ContentLength));
                st.Close();
                st.Dispose();
            }

            return bytes;

            //WebRequest request = WebRequest.Create(url);
            //WebResponse response = request.GetResponse();
            //Stream s = response.GetResponseStream();
            //byte[] data = new byte[response.ContentLength];
            //int length = 0;
            //MemoryStream ms = new MemoryStream();
            //while ((length = s.Read(data, 0, data.Length)) > 0)
            //{
            //    ms.Write(data, 0, length);
            //}
            //ms.Seek(0, SeekOrigin.Begin);
            //return data;
        }
    }

}
