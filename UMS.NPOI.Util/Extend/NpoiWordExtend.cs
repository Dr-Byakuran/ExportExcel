﻿using System;
using System.Drawing;
using System.IO;
using System.Net;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;

/********************************************************************************
 ** 版 本：
 ** Copyright (c) 2015-2018 厦门攸信信息技术有限公司
 ** 创 建：詹建妹（james_zhan@intretech.com）
 ** 日 期：2019/01/15 17:04:00
 ** 描 述：
*********************************************************************************/
namespace UMS.Framework.NpoiUtil
{
    /// <summary>
    /// Npoi Word 操作扩展
    /// </summary>
    public static class NpoiWordExtend
    {
        /// <summary>
        /// 添加图片
        /// </summary>
        /// <param name="dc"></param>
        /// <param name="path"></param>
        public static void AddPicture(this XWPFDocument doc, string path, AddPictureType pType)
        {
            if(path.IndexOf("http") != -1)
            {
                //byte[] bytes = ExcelUtil.getFile(path);
                WebRequest request = WebRequest.Create(path);
                WebResponse response = request.GetResponse();
                Stream s = response.GetResponseStream();
                byte[] data = new byte[response.ContentLength];
                int length = 0;
                MemoryStream ms = new MemoryStream();
                while ((length = s.Read(data, 0, data.Length)) > 0)
                {
                    ms.Write(data, 0, length);
                }
                ms.Seek(0, SeekOrigin.Begin);

                //using (Stream stt = new MemoryStream(bytes))
                //{
                    CT_P p = doc.GetNewP();
                    p.SetAlign(ST_Jc.center);
                    XWPFParagraph gp = new XWPFParagraph(p, doc);
                    XWPFRun run = gp.CreateRun();
                    run.AddPicture(ms, (int)PictureType.PNG, "2.png", 1000000, 1000000);
                ms.Close();
                    //stt.Close();
                //}
            }
            //else if(path.IndexOf(";base64") != -1)
            //{
            //    int index = path.IndexOf(",");
            //    path = path.Substring(index + 1);
            //    byte[] bytes = Convert.FromBase64String(path);
            //    using (Stream s64 = new MemoryStream(bytes))
            //    {
            //        CT_P p = doc.GetNewP();
            //        p.SetAlign(ST_Jc.center);
            //        XWPFParagraph gp = new XWPFParagraph(p, doc);
            //        XWPFRun run = gp.CreateRun();
            //        run.AddPicture(s64, (int)PictureType.PNG, "1.png", 1000000, 1000000);
            //        s64.Close();
            //    }
            //}
            //else
            //{
            //    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            //    {
            //        CT_P p = doc.GetNewP();
            //        p.SetAlign(ST_Jc.center);
            //        XWPFParagraph gp = new XWPFParagraph(p, doc);
            //        XWPFRun run = gp.CreateRun();
            //        run.AddPicture(fs, (int)PictureType.PNG, "1.png", 1000000, 1000000);
            //        fs.Close();
            //    }
            //}
        }

        /// <summary>
        /// 获取缩进值
        /// </summary>
        /// <param name="nameType">字体名称</param>
        /// <param name="sizeType">字号大小</param>
        /// <param name="indentationPoints">缩进字符数</param>
        public static int GetIndentation(this XWPFParagraph pg,FontNameType nameType, FontSizeType sizeType, int indentationPoints)
        {
            string fontName = GetFontName(nameType);
            int fontSize = (int)GetFontSize(sizeType);
            int len = GetIndentation(fontName, fontSize, indentationPoints, FontStyle.Regular);
            return len;
        }

        /// <summary>
        /// 获取缩进值
        /// </summary>
        /// <param name="pg"></param>
        /// <param name="indentationPoints"></param>
        public static int GetIndentation(this XWPFParagraph pg, int indentationPoints)
        {
            int len = 0;
            if(pg.Runs.Count > 0)
            {
                XWPFRun run = pg.Runs[0];
                len = GetIndentation(run.FontFamily, run.FontSize * 2, indentationPoints, FontStyle.Regular);
            }
            return len;
        }

        /// <summary>
        /// 获取一个新文档
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static CT_SectPr GetNewSectPr(this XWPFDocument doc)
        {
            doc.Document.body.sectPr = new CT_SectPr();
            CT_SectPr sectPr = doc.Document.body.sectPr;
            return sectPr;
        }

        /// <summary>
        /// 获取一个新段落
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static CT_P GetNewP(this XWPFDocument doc)
        {
            return doc.Document.body.AddNewP();
        }

        /// <summary>
        /// 设置段落对齐方式
        /// </summary>
        /// <param name="p"></param>
        /// <param name="align"></param>
        public static void SetAlign(this CT_P p, ST_Jc align)
        {
            CT_PPr ppr = p.pPr == null ? p.AddNewPPr() : p.pPr;
            CT_Jc jc = ppr.IsSetJc() ? ppr.jc : ppr.AddNewJc();
            jc.val = align;
        }

        /// <summary>
        /// 设置页面尺寸
        /// </summary>
        /// <param name="setPr"></param>
        /// <param name="pageType"></param>
        public static void SetPageSize(this CT_SectPr setPr, PageType pageType)
        {
            Tuple<ulong, ulong> size = PageSize(pageType);
            setPr.pgSz.w = size.Item1;
            setPr.pgSz.h = size.Item2;
        }

        /// <summary>
        /// 设置颜色
        /// </summary>
        /// <param name="run"></param>
        /// <param name="colorType"></param>
        public static void SetRunColor(this XWPFRun run, ColorType colorType)
        {
            Tuple<string, string> temp = GetColor(colorType);
            CT_R r = run.GetCTR();
            CT_RPr rpr = r.IsSetRPr() ? r.rPr : r.AddNewRPr();
            CT_Color color = rpr.IsSetColor() ? rpr.color : rpr.AddNewColor();
            color.val = temp.Item1;
        }

        /// <summary>
        /// 获取颜色 16进制色值
        /// </summary>
        /// <param name="run"></param>
        /// <param name="colorType"></param>
        /// <returns></returns>
        public static string GetColorHex(ColorType colorType)
        {
            Tuple<string, string> color = GetColor(colorType);
            return color.Item1;
        }

        /// <summary>
        /// 设置字体大小
        /// </summary>
        /// <param name="run"></param>
        /// <param name="sizeType"></param>
        public static void SetFontSize(this XWPFRun run, FontSizeType sizeType)
        {
            ulong fontSize = GetFontSize(sizeType);
            CT_R r = run.GetCTR();
            CT_RPr rpr = r.IsSetRPr() ? r.rPr : r.AddNewRPr();
            CT_HpsMeasure ctSize = rpr.IsSetSz() ? rpr.sz : rpr.AddNewSz();
            ctSize.val = fontSize;
        }

        /// <summary>
        /// 设置字体
        /// </summary>
        /// <param name="run"></param>
        /// <param name="nameType"></param>
        public static void SetFontName(this XWPFRun run, FontNameType nameType)
        {
            string fontName = GetFontName(nameType);
            CT_R r = run.GetCTR();
            CT_RPr rpr = r.IsSetRPr() ? r.rPr : r.AddNewRPr();
            CT_Fonts fonts = rpr.IsSetRFonts() ? rpr.rFonts : rpr.AddNewRFonts();
            fonts.ascii = fontName;
            if (string.IsNullOrEmpty(fonts.eastAsia))
                fonts.eastAsia = fontName;
            if (string.IsNullOrEmpty(fonts.cs))
                fonts.cs = fontName;
            if (string.IsNullOrEmpty(fonts.hAnsi))
                fonts.hAnsi = fontName;
        }

        /// <summary>
        /// 获取字体大小
        /// </summary>
        /// <param name="sizeType"></param>
        /// <returns></returns>
        private static ulong GetFontSize(FontSizeType sizeType)
        {
            double fontSize = 12;
            switch (sizeType)
            {
                case FontSizeType.初号:
                    fontSize = 42;
                    break;
                case FontSizeType.小初:
                    fontSize = 36;
                    break;
                case FontSizeType.一号:
                    fontSize = 36;
                    break;
                case FontSizeType.小一:
                    fontSize = 24;
                    break;
                case FontSizeType.二号:
                    fontSize = 22;
                    break;
                case FontSizeType.小二:
                    fontSize = 18;
                    break;
                case FontSizeType.三号:
                    fontSize = 16;
                    break;
                case FontSizeType.小三:
                    fontSize = 15;
                    break;
                case FontSizeType.四号:
                    fontSize = 14;
                    break;
                case FontSizeType.小四:
                    fontSize = 12;
                    break;
                case FontSizeType.五号:
                    fontSize = 10.5;
                    break;
                case FontSizeType.小五:
                    fontSize = 9;
                    break;
                case FontSizeType.六号:
                    fontSize = 7.5;
                    break;
                case FontSizeType.小六:
                    fontSize = 6.5;
                    break;
                case FontSizeType.七号:
                    fontSize = 5.5;
                    break;
                case FontSizeType.八号:
                    fontSize = 5;
                    break;
                case FontSizeType.Num5:
                    fontSize = 5;
                    break;
                case FontSizeType.Num5p5:
                    fontSize = 5.5;
                    break;
                case FontSizeType.Num6p5:
                    fontSize = 6.5;
                    break;
                case FontSizeType.Num7p5:
                    fontSize = 7.5;
                    break;
                case FontSizeType.Num8:
                    fontSize = 8;
                    break;
                case FontSizeType.Num9:
                    fontSize = 9;
                    break;
                case FontSizeType.Num10p5:
                    fontSize = 10.5;
                    break;
                case FontSizeType.Num11:
                    fontSize = 11;
                    break;
                case FontSizeType.Num12:
                    fontSize = 12;
                    break;
                case FontSizeType.Num14:
                    fontSize = 14;
                    break;
                case FontSizeType.Num16:
                    fontSize = 16;
                    break;
                case FontSizeType.Num18:
                    fontSize = 18;
                    break;
                case FontSizeType.Num20:
                    fontSize = 20;
                    break;
                case FontSizeType.Num22:
                    fontSize = 22;
                    break;
                case FontSizeType.Num24:
                    fontSize = 24;
                    break;
                case FontSizeType.Num26:
                    fontSize = 26;
                    break;
                case FontSizeType.Num28:
                    fontSize = 28;
                    break;
                case FontSizeType.Num36:
                    fontSize = 36;
                    break;
                case FontSizeType.Num48:
                    fontSize = 48;
                    break;
                case FontSizeType.Num72:
                    fontSize = 72;
                    break;
            }
            return  (ulong)fontSize * 2;
        }

        /// <summary>
        /// 获取纸张大小; Item1 宽 Item2 高； 默认 A4纸张
        /// <para>换算关系：1英寸=1440缇  1厘米=567缇  1磅=20缇  1像素=15缇</para>
        /// </summary>
        /// <param name="pageType"></param>
        /// <returns></returns>
        private static Tuple<ulong, ulong> PageSize(PageType pageType)
        {
            Tuple<ulong, ulong> size = null;
            switch (pageType)
            {
                case PageType.LetterPager_W:
                    size = new Tuple<ulong, ulong>(12241, 15841);
                    break;
                case PageType.LetterPager_H:
                    size = new Tuple<ulong, ulong>(15841, 12241);
                    break;
                case PageType.A4_W:
                    size = new Tuple<ulong, ulong>(11907, 16839);
                    break;
                case PageType.A4_H:
                    size = new Tuple<ulong, ulong>(16839, 11907);
                    break;
                case PageType.K16_W:
                    size = new Tuple<ulong, ulong>(10432, 14742);
                    break;
                case PageType.K16_H:
                    size = new Tuple<ulong, ulong>(14742, 10432);
                    break;
                case PageType.K32_W:
                    size = new Tuple<ulong, ulong>(7371, 10432);
                    break;
                case PageType.K32_H:
                    size = new Tuple<ulong, ulong>(10432, 7371);
                    break;
                case PageType.KB32_W:
                    size = new Tuple<ulong, ulong>(7938, 11510);
                    break;
                case PageType.KB32_H:
                    size = new Tuple<ulong, ulong>(11510, 7938);
                    break;
                default:
                    size = new Tuple<ulong, ulong>(11907, 16839);
                    break;
            }
            return size;
        }

        /// <summary>
        /// 返回设置颜色
        /// </summary>
        /// <param name="colorType"></param>
        /// <returns></returns>
        private static Tuple<string,string> GetColor(ColorType colorType)
        {
            Tuple<string, string> color = null;
            switch (colorType)
            {

                case ColorType.aliceblue:
                    color = new Tuple<string, string>("f0f8ff", "240,248,255");
                    break;
                case ColorType.antiquewhite:
                    color = new Tuple<string, string>("faebd7", "250,235,215");
                    break;
                case ColorType.aqua:
                    color = new Tuple<string, string>("00ffff", "0,255,255");
                    break;
                case ColorType.aquamarine:
                    color = new Tuple<string, string>("7fffd4", "127,255,212");
                    break;
                case ColorType.azure:
                    color = new Tuple<string, string>("f0ffff", "240,255,255");
                    break;
                case ColorType.beige:
                    color = new Tuple<string, string>("f5f5dc", "245,245,220");
                    break;
                case ColorType.bisque:
                    color = new Tuple<string, string>("ffe4c4", "255,228,196");
                    break;
                case ColorType.black:
                    color = new Tuple<string, string>("000000", "0,0,0");
                    break;
                case ColorType.blanchedalmond:
                    color = new Tuple<string, string>("ffebcd", "255,235,205");
                    break;
                case ColorType.blue:
                    color = new Tuple<string, string>("0000ff", "0,0,255");
                    break;
                case ColorType.blueviolet:
                    color = new Tuple<string, string>("8a2be2", "138,43,226");
                    break;
                case ColorType.brown:
                    color = new Tuple<string, string>("a52a2a", "165,42,42");
                    break;
                case ColorType.burlywood:
                    color = new Tuple<string, string>("deb887", "222,184,135");
                    break;
                case ColorType.cadetblue:
                    color = new Tuple<string, string>("5f9ea0", "95,158,160");
                    break;
                case ColorType.chartreuse:
                    color = new Tuple<string, string>("7fff00", "127,255,0");
                    break;
                case ColorType.chocolate:
                    color = new Tuple<string, string>("d2691e", "210,105,30");
                    break;
                case ColorType.coral:
                    color = new Tuple<string, string>("ff7f50", "255,127,80");
                    break;
                case ColorType.cornflowerblue:
                    color = new Tuple<string, string>("6495ed", "100,149,237");
                    break;
                case ColorType.cornsilk:
                    color = new Tuple<string, string>("fff8dc", "255,248,220");
                    break;
                case ColorType.crimson:
                    color = new Tuple<string, string>("dc143c", "220,20,60");
                    break;
                case ColorType.cyan:
                    color = new Tuple<string, string>("00ffff", "0,255,255");
                    break;
                case ColorType.darkblue:
                    color = new Tuple<string, string>("00008b", "0,0,139");
                    break;
                case ColorType.darkcyan:
                    color = new Tuple<string, string>("008b8b", "0,139,139");
                    break;
                case ColorType.darkgoldenrod:
                    color = new Tuple<string, string>("b8860b", "184,134,11");
                    break;
                case ColorType.darkgray:
                    color = new Tuple<string, string>("a9a9a9", "169,169,169");
                    break;
                case ColorType.darkgreen:
                    color = new Tuple<string, string>("006400", "0,100,0");
                    break;
                case ColorType.darkgrey:
                    color = new Tuple<string, string>("a9a9a9", "169,169,169");
                    break;
                case ColorType.darkkhaki:
                    color = new Tuple<string, string>("bdb76b", "189,183,107");
                    break;
                case ColorType.darkmagenta:
                    color = new Tuple<string, string>("8b008b", "139,0,139");
                    break;
                case ColorType.darkolivegreen:
                    color = new Tuple<string, string>("556b2f", "85,107,47");
                    break;
                case ColorType.darkorange:
                    color = new Tuple<string, string>("ff8c00", "255,140,0");
                    break;
                case ColorType.darkorchid:
                    color = new Tuple<string, string>("9932cc", "153,50,204");
                    break;
                case ColorType.darkred:
                    color = new Tuple<string, string>("8b0000", "139,0,0");
                    break;
                case ColorType.darksalmon:
                    color = new Tuple<string, string>("e9967a", "233,150,122");
                    break;
                case ColorType.darkseagreen:
                    color = new Tuple<string, string>("8fbc8f", "143,188,143");
                    break;
                case ColorType.darkslateblue:
                    color = new Tuple<string, string>("483d8b", "72,61,139");
                    break;
                case ColorType.darkslategray:
                    color = new Tuple<string, string>("2f4f4f", "47,79,79");
                    break;
                case ColorType.darkslategrey:
                    color = new Tuple<string, string>("2f4f4f", "47,79,79");
                    break;
                case ColorType.darkturquoise:
                    color = new Tuple<string, string>("00ced1", "0,206,209");
                    break;
                case ColorType.darkviolet:
                    color = new Tuple<string, string>("9400d3", "148,0,211");
                    break;
                case ColorType.deeppink:
                    color = new Tuple<string, string>("ff1493", "255,20,147");
                    break;
                case ColorType.deepskyblue:
                    color = new Tuple<string, string>("00bfff", "0,191,255");
                    break;
                case ColorType.dimgray:
                    color = new Tuple<string, string>("696969", "105,105,105");
                    break;
                case ColorType.dimgrey:
                    color = new Tuple<string, string>("696969", "105,105,105");
                    break;
                case ColorType.dodgerblue:
                    color = new Tuple<string, string>("1e90ff", "30,144,255");
                    break;
                case ColorType.firebrick:
                    color = new Tuple<string, string>("b22222", "178,34,34");
                    break;
                case ColorType.floralwhite:
                    color = new Tuple<string, string>("fffaf0", "255,250,240");
                    break;
                case ColorType.forestgreen:
                    color = new Tuple<string, string>("228b22", "34,139,34");
                    break;
                case ColorType.fuchsia:
                    color = new Tuple<string, string>("ff00ff", "255,0,255");
                    break;
                case ColorType.gainsboro:
                    color = new Tuple<string, string>("dcdcdc", "220,220,220");
                    break;
                case ColorType.ghostwhite:
                    color = new Tuple<string, string>("f8f8ff", "248,248,255");
                    break;
                case ColorType.gold:
                    color = new Tuple<string, string>("ffd700", "255,215,0");
                    break;
                case ColorType.goldenrod:
                    color = new Tuple<string, string>("daa520", "218,165,32");
                    break;
                case ColorType.gray:
                    color = new Tuple<string, string>("808080", "128,128,128");
                    break;
                case ColorType.green:
                    color = new Tuple<string, string>("008000", "0,128,0");
                    break;
                case ColorType.greenyellow:
                    color = new Tuple<string, string>("adff2f", "173,255,47");
                    break;
                case ColorType.grey:
                    color = new Tuple<string, string>("808080", "128,128,128");
                    break;
                case ColorType.honeydew:
                    color = new Tuple<string, string>("f0fff0", "240,255,240");
                    break;
                case ColorType.hotpink:
                    color = new Tuple<string, string>("ff69b4", "255,105,180");
                    break;
                case ColorType.indianred:
                    color = new Tuple<string, string>("cd5c5c", "205,92,92");
                    break;
                case ColorType.indigo:
                    color = new Tuple<string, string>("4b0082", "75,0,130");
                    break;
                case ColorType.ivory:
                    color = new Tuple<string, string>("fffff0", "255,255,240");
                    break;
                case ColorType.khaki:
                    color = new Tuple<string, string>("f0e68c", "240,230,140");
                    break;
                case ColorType.lavender:
                    color = new Tuple<string, string>("e6e6fa", "230,230,250");
                    break;
                case ColorType.lavenderblush:
                    color = new Tuple<string, string>("fff0f5", "255,240,245");
                    break;
                case ColorType.lawngreen:
                    color = new Tuple<string, string>("7cfc00", "124,252,0");
                    break;
                case ColorType.lemonchiffon:
                    color = new Tuple<string, string>("fffacd", "255,250,205");
                    break;
                case ColorType.lightblue:
                    color = new Tuple<string, string>("add8e6", "173,216,230");
                    break;
                case ColorType.lightcoral:
                    color = new Tuple<string, string>("f08080", "240,128,128");
                    break;
                case ColorType.lightcyan:
                    color = new Tuple<string, string>("e0ffff", "224,255,255");
                    break;
                case ColorType.lightgoldenrodyellow:
                    color = new Tuple<string, string>("fafad2", "250,250,210");
                    break;
                case ColorType.lightgray:
                    color = new Tuple<string, string>("d3d3d3", "211,211,211");
                    break;
                case ColorType.lightgreen:
                    color = new Tuple<string, string>("90ee90", "144,238,144");
                    break;
                case ColorType.lightgrey:
                    color = new Tuple<string, string>("d3d3d3", "211,211,211");
                    break;
                case ColorType.lightpink:
                    color = new Tuple<string, string>("ffb6c1", "255,182,193");
                    break;
                case ColorType.lightsalmon:
                    color = new Tuple<string, string>("ffa07a", "255,160,122");
                    break;
                case ColorType.lightseagreen:
                    color = new Tuple<string, string>("20b2aa", "32,178,170");
                    break;
                case ColorType.lightskyblue:
                    color = new Tuple<string, string>("87cefa", "135,206,250");
                    break;
                case ColorType.lightslategray:
                    color = new Tuple<string, string>("778899", "119,136,153");
                    break;
                case ColorType.lightslategrey:
                    color = new Tuple<string, string>("778899", "119,136,153");
                    break;
                case ColorType.lightsteelblue:
                    color = new Tuple<string, string>("b0c4de", "176,196,222");
                    break;
                case ColorType.lightyellow:
                    color = new Tuple<string, string>("ffffe0", "255,255,224");
                    break;
                case ColorType.lime:
                    color = new Tuple<string, string>("00ff00", "0,255,0");
                    break;
                case ColorType.limegreen:
                    color = new Tuple<string, string>("32cd32", "50,205,50");
                    break;
                case ColorType.linen:
                    color = new Tuple<string, string>("faf0e6", "250,240,230");
                    break;
                case ColorType.magenta:
                    color = new Tuple<string, string>("ff00ff", "255,0,255");
                    break;
                case ColorType.maroon:
                    color = new Tuple<string, string>("800000", "128,0,0");
                    break;
                case ColorType.mediumaquamarine:
                    color = new Tuple<string, string>("66cdaa", "102,205,170");
                    break;
                case ColorType.mediumblue:
                    color = new Tuple<string, string>("0000cd", "0,0,205");
                    break;
                case ColorType.mediumorchid:
                    color = new Tuple<string, string>("ba55d3", "186,85,211");
                    break;
                case ColorType.mediumpurple:
                    color = new Tuple<string, string>("9370db", "147,112,219");
                    break;
                case ColorType.mediumseagreen:
                    color = new Tuple<string, string>("3cb371", "60,179,113");
                    break;
                case ColorType.mediumslateblue:
                    color = new Tuple<string, string>("7b68ee", "123,104,238");
                    break;
                case ColorType.mediumspringgreen:
                    color = new Tuple<string, string>("00fa9a", "0,250,154");
                    break;
                case ColorType.mediumturquoise:
                    color = new Tuple<string, string>("48d1cc", "72,209,204");
                    break;
                case ColorType.mediumvioletred:
                    color = new Tuple<string, string>("c71585", "199,21,133");
                    break;
                case ColorType.midnightblue:
                    color = new Tuple<string, string>("191970", "25,25,112");
                    break;
                case ColorType.mintcream:
                    color = new Tuple<string, string>("f5fffa", "245,255,250");
                    break;
                case ColorType.mistyrose:
                    color = new Tuple<string, string>("ffe4e1", "255,228,225");
                    break;
                case ColorType.moccasin:
                    color = new Tuple<string, string>("ffe4b5", "255,228,181");
                    break;
                case ColorType.navajowhite:
                    color = new Tuple<string, string>("ffdead", "255,222,173");
                    break;
                case ColorType.navy:
                    color = new Tuple<string, string>("000080", "0,0,128");
                    break;
                case ColorType.oldlace:
                    color = new Tuple<string, string>("fdf5e6", "253,245,230");
                    break;
                case ColorType.olive:
                    color = new Tuple<string, string>("808000", "128,128,0");
                    break;
                case ColorType.olivedrab:
                    color = new Tuple<string, string>("6b8e23", "107,142,35");
                    break;
                case ColorType.orange:
                    color = new Tuple<string, string>("ffa500", "255,165,0");
                    break;
                case ColorType.orangered:
                    color = new Tuple<string, string>("ff4500", "255,69,0");
                    break;
                case ColorType.orchid:
                    color = new Tuple<string, string>("da70d6", "218,112,214");
                    break;
                case ColorType.palegoldenrod:
                    color = new Tuple<string, string>("eee8aa", "238,232,170");
                    break;
                case ColorType.palegreen:
                    color = new Tuple<string, string>("98fb98", "152,251,152");
                    break;
                case ColorType.paleturquoise:
                    color = new Tuple<string, string>("afeeee", "175,238,238");
                    break;
                case ColorType.palevioletred:
                    color = new Tuple<string, string>("db7093", "219,112,147");
                    break;
                case ColorType.papayawhip:
                    color = new Tuple<string, string>("ffefd5", "255,239,213");
                    break;
                case ColorType.peachpuff:
                    color = new Tuple<string, string>("ffdab9", "255,218,185");
                    break;
                case ColorType.peru:
                    color = new Tuple<string, string>("cd853f", "205,133,63");
                    break;
                case ColorType.pink:
                    color = new Tuple<string, string>("ffc0cb", "255,192,203");
                    break;
                case ColorType.plum:
                    color = new Tuple<string, string>("dda0dd", "221,160,221");
                    break;
                case ColorType.powderblue:
                    color = new Tuple<string, string>("b0e0e6", "176,224,230");
                    break;
                case ColorType.purple:
                    color = new Tuple<string, string>("800080", "128,0,128");
                    break;
                case ColorType.red:
                    color = new Tuple<string, string>("ff0000", "255,0,0");
                    break;
                case ColorType.rosybrown:
                    color = new Tuple<string, string>("bc8f8f", "188,143,143");
                    break;
                case ColorType.royalblue:
                    color = new Tuple<string, string>("4169e1", "65,105,225");
                    break;
                case ColorType.saddlebrown:
                    color = new Tuple<string, string>("8b4513", "139,69,19");
                    break;
                case ColorType.salmon:
                    color = new Tuple<string, string>("fa8072", "250,128,114");
                    break;
                case ColorType.sandybrown:
                    color = new Tuple<string, string>("f4a460", "244,164,96");
                    break;
                case ColorType.seagreen:
                    color = new Tuple<string, string>("2e8b57", "46,139,87");
                    break;
                case ColorType.seashell:
                    color = new Tuple<string, string>("fff5ee", "255,245,238");
                    break;
                case ColorType.sienna:
                    color = new Tuple<string, string>("a0522d", "160,82,45");
                    break;
                case ColorType.silver:
                    color = new Tuple<string, string>("c0c0c0", "192,192,192");
                    break;
                case ColorType.skyblue:
                    color = new Tuple<string, string>("87ceeb", "135,206,235");
                    break;
                case ColorType.slateblue:
                    color = new Tuple<string, string>("6a5acd", "106,90,205");
                    break;
                case ColorType.slategray:
                    color = new Tuple<string, string>("708090", "112,128,144");
                    break;
                case ColorType.slategrey:
                    color = new Tuple<string, string>("708090", "112,128,144");
                    break;
                case ColorType.snow:
                    color = new Tuple<string, string>("fffafa", "255,250,250");
                    break;
                case ColorType.springgreen:
                    color = new Tuple<string, string>("00ff7f", "0,255,127");
                    break;
                case ColorType.steelblue:
                    color = new Tuple<string, string>("4682b4", "70,130,180");
                    break;
                case ColorType.tan:
                    color = new Tuple<string, string>("d2b48c", "210,180,140");
                    break;
                case ColorType.teal:
                    color = new Tuple<string, string>("008080", "0,128,128");
                    break;
                case ColorType.thistle:
                    color = new Tuple<string, string>("d8bfd8", "216,191,216");
                    break;
                case ColorType.tomato:
                    color = new Tuple<string, string>("ff6347", "255,99,71");
                    break;
                case ColorType.turquoise:
                    color = new Tuple<string, string>("40e0d0", "64,224,208");
                    break;
                case ColorType.violet:
                    color = new Tuple<string, string>("ee82ee", "238,130,238");
                    break;
                case ColorType.wheat:
                    color = new Tuple<string, string>("f5deb3", "245,222,179");
                    break;
                case ColorType.white:
                    color = new Tuple<string, string>("ffffff", "255,255,255");
                    break;
                case ColorType.whitesmoke:
                    color = new Tuple<string, string>("f5f5f5", "245,245,245");
                    break;
                case ColorType.yellow:
                    color = new Tuple<string, string>("ffff00", "255,255,0");
                    break;
                case ColorType.yellowgreen:
                    color = new Tuple<string, string>("9acd32", "154,205,50");
                    break;
            }

            return color;
        }

        /// <summary>
        /// 获取字体名称
        /// </summary>
        /// <param name="nameType"></param>
        /// <returns></returns>
        private static string GetFontName(FontNameType nameType)
        {
            string fontName = string.Empty;
            switch (nameType)
            {
                case FontNameType.MingLiU_HKSCS_ExtB:
                    fontName = "MingLiU_HKSCS-ExtB";
                        break;
                case FontNameType.MingLiU_ExtB:
                    fontName = "MingLiU-ExtB";
                    break;
                case FontNameType.PMingLiU_ExtB:
                    fontName = "PMingLiU-ExtB";
                    break;
                case FontNameType.SimSun_ExtB:
                    fontName = "SimSun-ExtB";
                    break;
                case FontNameType.Modern_No20:
                    fontName = "Modern No.20";
                    break;
                default:
                    fontName = nameType.ToString().Replace("_", " ");
                    break;
            }
            if (string.IsNullOrEmpty(fontName))
                fontName = FontNameType.宋体.ToString();
            return fontName;
        }

        /// <summary>
        /// 获取缩进点数
        /// </summary>
        /// <param name="fontname"></param>
        /// <param name="fontsize"></param>
        /// <param name="Indentationfonts"></param>
        /// <param name="fs"></param>
        /// <returns></returns>
        private static int GetIndentation(string fontName, int fontSize, int IndentationFonts, FontStyle fs)
        {
            Pen pen = new Pen(Color.White);
            Bitmap bm = new Bitmap(50, 100);
            Graphics m_tmpGr = Graphics.FromImage(bm);

            m_tmpGr.PageUnit = GraphicsUnit.Point;
            SizeF size = m_tmpGr.MeasureString("人", new Font(fontName, fontSize * 0.75F, fs));
            return (int)size.Width * IndentationFonts * 10;
        }
    }
}
