using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using UMS.Framework.NpoiUtil.Model;
using UMS.Framework.NpoiUtil.Util;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil
{
    public class TemplateDownloadUtil
    {
        XWPFDocument doc = null;
        List<XWPFParagraph> listParagraph = null;
        List<TableCell> listCell = null;
        private string validSingle = "$";
        private string validDouble = "&";
        private string matchSingle = "\\$\\{(.+?)\\}";
        private string repSingle = "[\\${\\}]";
        private string matchDouble = "\\&\\{(.+?)\\}";
        private string repDouble = "[&{\\}]";

        public TemplateDownloadUtil()
        {
            listParagraph = new List<XWPFParagraph>();
            listCell = new List<TableCell>();
        }

        /// <summary>
        /// 下载Word  
        /// XWPFRun表示有相同属性的一段文本，所以模板里变量内容需要从左到右的顺序写，${userName}，如果先写${},再添加内容，会拆分成几部分，不能正常使用
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <param name="entity"></param>
        /// <returns></returns>
        public string DownLoad(string path, string fileName, DynamicEntity entity)
        {
            string errorMessage = string.Empty;
            errorMessage = BeforeDownload(path);
            if (!string.IsNullOrEmpty(errorMessage)) return errorMessage;
            if (listParagraph == null || listParagraph.Any()) return "模板为空";
            foreach (var item in listParagraph)
            {
                ProcessParagraph(item, entity);
            }

            HttpWrite(fileName);
            return errorMessage;
        }

        /// <summary>
        /// 下载Word
        /// XWPFRun表示有相同属性的一段文本，所以模板里变量内容需要从左到右的顺序写，${userName}，如果先写${},再添加内容，会拆分成几部分，不能正常使用
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        public string DownLoad(string path, string fileName, IEnumerable<DynamicEntity> list)
        {
            string errorMessage = string.Empty;
            errorMessage = BeforeDownload(path);
            if (string.IsNullOrEmpty(errorMessage) == false) return errorMessage;
            if (listCell == null || listCell.Any() == false) return "模板为空";
            foreach (var tableCell in listCell)
            {
                ProcessTableCell(tableCell, list);
            }

            HttpWrite(fileName);
            return errorMessage;
        }

        /// <summary>
        /// 下载 Word
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <param name="main"></param>
        /// <param name="sub"></param>
        /// <returns></returns>
        public string DownLoad(string path, string fileName, DynamicEntity main, IEnumerable<DynamicEntity> sub)
        {
            string errorMessage = string.Empty;
            errorMessage = BeforeDownload(path);
            if (string.IsNullOrEmpty(errorMessage) == false) return errorMessage;
            if ((listParagraph == null || listParagraph.Any() == false) && main != null) return "模板为空";
            if (listCell == null || listCell.Any() == false) return "模板为空";
            foreach (var para in listParagraph)
            {
                ProcessParagraph(para, main);
            }
            foreach (var tableCell in listCell)
            {
                ProcessTableCell(tableCell, sub);
            }

            HttpWrite(fileName);
            return errorMessage;
        }

        public string DownLoad(string path, string fileName, DynamicEntity main, params IEnumerable<DynamicEntity>[] list)
        {
            string errorMessage = string.Empty;
            errorMessage = BeforeDownload(path);
            if (string.IsNullOrEmpty(errorMessage) == false) return errorMessage;
            if (listParagraph == null || listParagraph.Any() == false) return "模板为空";
            if (listCell == null || listCell.Any() == false) return "模板为空";
            foreach (var param in listParagraph)
            {
                ProcessParagraph(param, main);
            }
            foreach(var tableCell in listCell)
            {
                foreach(var sub in list)
                {
                    ProcessTableCell(tableCell, sub);
                }
            }
            HttpWrite(fileName);
            return errorMessage;
        }

        #region 私有处理方法

        /// <summary>
        /// 在下载之前
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private string BeforeDownload(string path, bool includeTable = false)
        {
            string errorMessage = string.Empty;
            try
            {
                byte[] bytes = GetHttpUrlData(path);
                using (MemoryStream ms = new MemoryStream(bytes,0,bytes.Length))
                {
                    doc = new XWPFDocument(ms);
                    foreach (XWPFParagraph item in doc.Paragraphs)
                    {
                        if (string.IsNullOrEmpty(item.ParagraphText)) continue;
                        if (item.ParagraphText.Contains(validSingle))
                        {
                            listParagraph.Add(item);
                        }
                    }
                    for (var tbIndex = 0; tbIndex < doc.Tables.Count; tbIndex++)
                    {
                        XWPFTable tb = doc.Tables[tbIndex];
                        for (var rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++)
                        {
                            XWPFTableRow row = tb.Rows[rowIndex];
                            List<XWPFTableCell> cells = row.GetTableCells();
                            for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                            {
                                XWPFTableCell cell = cells[cellIndex];
                                foreach (XWPFParagraph item in cell.Paragraphs)
                                {
                                    if (string.IsNullOrEmpty(item.ParagraphText)) continue;
                                    if (item.ParagraphText.Contains(validDouble))
                                    {
                                        listCell.Add(new TableCell()
                                        {
                                            Table = tb,
                                            TableIndex = tbIndex,
                                            RowIndex = rowIndex,
                                            CellIndex = cellIndex,
                                            Cell = cell,
                                            Paragraph = item,
                                            CellCount = cells.Count
                                        });
                                    }
                                    if (item.ParagraphText.Contains(validSingle))
                                    {
                                        listParagraph.Add(item);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            return errorMessage;
        }

        /// <summary>
        /// 处理段落
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="entity"></param>
        private void ProcessParagraph(XWPFParagraph paragraph, DynamicEntity entity, bool isTableCell = false)
        {
            if (entity == null) return;
            Regex reg = null;
            string valid = string.Empty;
            if (isTableCell == false)
            {
                reg = new Regex(matchSingle, RegexOptions.Multiline | RegexOptions.Singleline);
                valid = validSingle;
            }
            else
            {
                reg = new Regex(matchDouble, RegexOptions.Multiline | RegexOptions.Singleline);
                valid = validDouble;
            }
            #region 废弃
            //for (int i = 0; i < paragraph.Runs.Count; i++)
            //{
            //    XWPFRun run = paragraph.Runs[i];
            //    string oldValue = run.Text;
            //    string newValue = run.Text;
            //    MatchCollection matchs = reg.Matches(run.Text);
            //    if (matchs == null || matchs.Count == 0) continue;
            //    foreach (Match match in matchs)
            //    {
            //        string propertyName = string.Empty;
            //        if (isTableCell == false)
            //            propertyName = Regex.Replace(match.Value, repSingle, "");
            //        else
            //            propertyName = Regex.Replace(match.Value, repDouble, "");
            //        object value = entity.GetPropertyValue(propertyName, false);
            //        //if (value == null) continue;
            //        if (value == null) value = string.Empty;
            //        newValue = newValue.Replace(match.Value, value.ToString());
            //    }
            //    var oldStyle = run.GetCTR().rPr;
            //    paragraph.RemoveRun(i);

            //    XWPFRun newRun = null;
            //    if (i == paragraph.Runs.Count && paragraph.Runs.Count > 2)
            //        newRun = paragraph.InsertNewRun(i - 1);
            //    else
            //        newRun = paragraph.InsertNewRun(i);
            //    newRun.SetText(newValue);
            //    newRun.GetCTR().rPr = oldStyle;
            //}
            #endregion
            MatchCollection matchs = reg.Matches(paragraph.Text);
            List<Tuple<XWPFRun,object, int>> newRuns = new List<Tuple<XWPFRun,object, int>>();
            foreach(Match match in matchs)
            {
                string propertyName = string.Empty;
                if (isTableCell)
                    propertyName = Regex.Replace(match.Value, repDouble, "");
                else
                    propertyName = Regex.Replace(match.Value, repSingle, "");
                var runs = paragraph.Runs.Where(t => t.Text.Contains(propertyName.Trim()));
                if(runs == null || runs.Any() == false)
                {
                    continue;
                }
                var run = runs.FirstOrDefault();
                int index = paragraph.Runs.IndexOf(run);
                int num = DealSurlusRun(paragraph, run, valid);
                //int num = 0;
                //if(index >= 1)
                //{
                //    var frontRun = paragraph.Runs[index - 1];
                //    if (frontRun.Text.Contains(valid))
                //    {
                //        paragraph.RemoveRun(index - 1);
                //        num += 1;
                //    }
                //}
                //var afterRun = paragraph.Runs[index + 1 - num];
                //if (afterRun.Text.TrimStart().StartsWith("}"))
                //{
                //    paragraph.RemoveRun(index + 1 - num);
                //}
                object value = entity.GetPropertyValue(propertyName, false);
                if (value == null) value = string.Empty;
                newRuns.Add(new Tuple<XWPFRun, object, int>(run, value, index - num));
            }
            foreach(var item in newRuns)
            {
                var oldStyle = item.Item1.GetCTR().rPr;
                //paragraph.RemoveRun(item.Item3);
                XWPFRun newRun = null;
                if (item.Item3 == paragraph.Runs.Count && paragraph.Runs.Count > 1)
                {
                    newRun = paragraph.InsertNewRun(item.Item3 - 1);
                }
                else
                {
                    if (item.Item3 == 0)
                    {
                        newRun = paragraph.InsertNewRun(0);
                        paragraph.RemoveRun(1);
                    }
                    else
                    {
                        paragraph.RemoveRun(item.Item3);
                        newRun = paragraph.InsertNewRun(item.Item3);
                    }
                }
                newRun.SetText(item.Item2.ToString());
                newRun.GetCTR().rPr = oldStyle;
            }
        }

        /// <summary>
        /// 处理表格
        /// </summary>
        /// <param name="tableCell"></param>
        /// <param name="list"></param>
        private void ProcessTableCell(TableCell tableCell, IEnumerable<DynamicEntity> list)
        {
            var newList = list.ToList();
            list = list.ToList();
            //if (tableCell.Paragraph.Runs.Count != 1) return;
            //Regex reg = new Regex(matchDouble, RegexOptions.Multiline | RegexOptions.Singleline);
            //MatchCollection matchs = reg.Matches(tableCell.Paragraph.ParagraphText);
            //if (matchs == null || matchs.Count != 1) return;
            //string propertyName = Regex.Replace(matchs[0].Value, repDouble, "");
            string propertyName = Regex.Replace(tableCell.Paragraph.Text, repDouble, "");
            var runs = tableCell.Paragraph.Runs.Where(t => t.Text.Contains(propertyName.Trim()));
            if (runs == null || runs.Any() == false)
            {
                return;
            }
            var run = runs.FirstOrDefault();
            int index = tableCell.Paragraph.Runs.IndexOf(run);
            CT_RPr oldStyle = tableCell.Paragraph.Runs[index].GetCTR().rPr;
            DealSurlusRun(tableCell.Paragraph, run, validDouble);
            index = tableCell.Paragraph.Runs.IndexOf(run);
            //int num = 0;
            //if (index >= 1)
            //{
            //    var frontRun = tableCell.Paragraph.Runs[index - 1];
            //    if (frontRun.Text.Contains(validDouble))
            //    {
            //        tableCell.Paragraph.RemoveRun(index - 1);
            //        num += 1;
            //    }
            //}
            //var afterRun = tableCell.Paragraph.Runs[index + 1 - num];
            //if (afterRun.Text.TrimStart().StartsWith("}"))
            //{
            //    tableCell.Paragraph.RemoveRun(index + 1 - num);
            //}
            int rowIndex = tableCell.RowIndex;
            var rowPr = tableCell.Table.GetRow(tableCell.RowIndex).GetCTRow().trPr;
            var cellPr = tableCell.Cell.GetCTTc().tcPr;
            for (var i = 0; i < list.Count(); i++)
            {
                DynamicEntity entity = newList[i];
                //if (entity.IsEntityProperty(propertyName.Trim()) == false) continue;
                object value = entity.GetPropertyValue(propertyName.Trim(), false);
                if (value == null) value = string.Empty;
                if (i == 0)
                {
                    tableCell.Paragraph.RemoveRun(index);
                    XWPFRun newRun = tableCell.Paragraph.CreateRun();
                    if (value != null)
                    {

                        if (value is byte[])
                        {
                            byte[] bytes = value as byte[];
                            using (MemoryStream ms = new MemoryStream(bytes, 0, bytes.Length))
                            {
                                newRun.AddPicture(ms, (int)PictureType.PNG, "test.png", NPOI.Util.Units.ToEMU(100), NPOI.Util.Units.ToEMU(100));
                                ms.Close();
                            }
                        }
                        else
                            newRun.SetText(value.ToString());
                    }
                    rowIndex += 1;
                    continue;
                }
                XWPFTableRow row = tableCell.Table.GetRow(rowIndex);
                if (row == null)
                {
                    row = tableCell.Table.CreateRow();
                    row.GetCTRow().trPr = rowPr;
                }
                XWPFTableCell cell = row.GetCell(tableCell.CellIndex);
                var cells = row.GetTableCells();
                if (cells != null && cells.Count == 1)
                {
                    string sdasd = string.Empty;
                    XWPFTableRow newRow = tableCell.Table.CreateRow();
                    newRow.GetCTRow().trPr = rowPr;
                    tableCell.Table.AddRow(newRow, rowIndex);
                    tableCell.Table.RemoveRow(rowIndex + 2);
                    cell = newRow.GetCell(tableCell.CellIndex);
                    newRow.GetCell(0).SetText(rowIndex.ToString());
                    newRow.GetCell(0).GetCTTc().AddNewTcPr();
                    newRow.GetCell(0).GetCTTc().tcPr = cellPr;               
                }
                if (cell == null) continue;
                if (value != null)
                {
                    //cell.SetText(value.ToString());
                    if (cell.Paragraphs == null || cell.Paragraphs.Count == 0)
                        cell.AddParagraph();
                    cell.Paragraphs[0].RemoveRun(0);
                    XWPFRun newRun = cell.Paragraphs[0].CreateRun();
                    if(value is byte[])
                    {
                        byte[] bytes = value as byte[];
                        using (MemoryStream ms = new MemoryStream(bytes, 0, bytes.Length))
                        {
                            newRun.AddPicture(ms, (int)PictureType.PNG, "test.png", NPOI.Util.Units.ToEMU(100), NPOI.Util.Units.ToEMU(100));
                            ms.Close();
                        }
                    }
                    else 
                        newRun.SetText(value.ToString());
                    newRun.GetCTR().rPr = oldStyle;
                    //XWPFRun newRun = cell.AddParagraph().CreateRun();
                    //newRun.SetText(value.ToString());
                    //newRun.GetCTR().rPr = oldStyle;
                }
                cell.GetCTTc().AddNewTcPr();
                cell.GetCTTc().tcPr = cellPr;
                rowIndex += 1;
            }
        }

        /// <summary>
        /// 写入Http
        /// </summary>
        /// <param name="fileName"></param>
        private void HttpWrite(string fileName)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Write(ms);
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearHeaders();
                HttpContext.Current.Response.Buffer = true;
                if (!StringUtil.Contains(HttpContext.Current.Request.UserAgent, "firefox", true) &&
                !StringUtil.Contains(HttpContext.Current.Request.UserAgent, "chrome", true))
                    fileName = StringUtil.UrlEncode(fileName, Encoding.UTF8, false);
                fileName = fileName.Replace("\"", "");
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + fileName);
                // 加入ContentType 防止火狐浏览器导出时直接导出Html，让其默认Excel导出
                //HttpContext.Current.Response.ContentType = "application/ms-excel";
                HttpContext.Current.Response.BinaryWrite(ms.ToArray());
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.End();
                ms.Close();
            }
        }

        private  byte[] GetHttpUrlData(string url)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Method = "GET";
            byte[] buffurPic = null;
            try
            {
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
                StreamReader reader = new StreamReader(webResponse.GetResponseStream(), Encoding.UTF8);
                Stream stream = webResponse.GetResponseStream();
                MemoryStream ms = null;
                Byte[] buffer = new Byte[webResponse.ContentLength];
                int offset = 0, actuallyRead = 0;
                do
                {
                    actuallyRead = stream.Read(buffer, offset, buffer.Length - offset);
                    offset += actuallyRead;
                }
                while (actuallyRead > 0);
                ms = new MemoryStream(buffer);

                buffurPic = ms.ToArray();
            }
            catch { }

            return buffurPic;
        }

        private int DealSurlusRun(XWPFParagraph paragraph, XWPFRun run, string valid)
        {
            int num = 0;
            int index = paragraph.Runs.IndexOf(run);
            if (index >= 1)
            {
                var frontRun = paragraph.Runs[index - 1];
                if (frontRun.Text.Contains(valid))
                {
                    paragraph.RemoveRun(index - 1);
                    num += 1;
                }
            }
            if (index + 1 - num < paragraph.Runs.Count)
            {
                var afterRun = paragraph.Runs[index + 1 - num];
                if (afterRun.Text.TrimStart().StartsWith("}"))
                {
                    paragraph.RemoveRun(index + 1 - num);
                }
            }
            return num;
        }
        #endregion
    }
}
