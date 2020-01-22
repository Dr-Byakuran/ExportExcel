using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UMS.Framework.NpoiUtil.Model;

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
    /// 颜色处理扩展
    /// </summary>
    public static class ExcelColorExtend
    {
        private static readonly short defaultColorIndexed = 9;
        private static List<Tuple<int, string>> originalRGBs = new List<Tuple<int, string>>();

        #region 自定义颜色

        /// <summary>
        /// 设置自定义RGB颜色
        /// </summary>
        /// <param name="workBook">工作簿</param>
        /// <param name="rgbs">RGB颜色集合</param>
        public static void SetCustomColor(this IWorkbook workBook, IEnumerable<string> rgbs)
        {
            if (workBook is HSSFWorkbook)
            {
                // 设置颜色
                var tempWork = (HSSFWorkbook)workBook;
                tempWork.SetCustomColor(rgbs);
            }
        }

        public static void SetCustomColor(this IWorkbook workBook, IEnumerable<ColorEntity> listColors)
        {
            if(workBook is HSSFWorkbook)
            {
                var tempWork = (HSSFWorkbook)workBook;

            }
        }

        /// <summary>
        /// 获取自定义颜色
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgb"></param>
        public static short GetCustomColor(this IWorkbook workBook, string rgb)
        {
            short indexed = defaultColorIndexed;
            if (workBook is HSSFWorkbook)
            {
                var tempWork = (HSSFWorkbook)workBook;
                indexed = tempWork.GetCustomColor(rgb);
            }
            else if (workBook is XSSFWorkbook)
            {
                var tempWork = (XSSFWorkbook)workBook;
                indexed = tempWork.GetCustomColor(rgb);
            }
            return indexed;
        }

        // 获取色值
        public static short GetCustomColor(this IWorkbook workBook, ColorType colorType)
        {
            short indexed = defaultColorIndexed;
            string rgb = ExcelExtend.GetColor(colorType).Item2;
            if (workBook is HSSFWorkbook)
            {
                var tempWork = (HSSFWorkbook)workBook;
                indexed = tempWork.GetCustomColor(rgb);
            }
            else if (workBook is XSSFWorkbook)
            {
                var tempWork = (XSSFWorkbook)workBook;
                indexed = tempWork.GetCustomColor(rgb);
            }
            return indexed;
        }

        /// <summary>
        /// 设置画板原生RGB颜色
        /// </summary>
        private static void SetOriginalRGB()
        {
            originalRGBs.Add(new Tuple<int, string>(8, "0,0,0"));
            originalRGBs.Add(new Tuple<int, string>(9, "255,255,255"));
            originalRGBs.Add(new Tuple<int, string>(10, "255,0,0"));
            originalRGBs.Add(new Tuple<int, string>(11, "0,255,0"));
            originalRGBs.Add(new Tuple<int, string>(12, "0,0,255"));
            originalRGBs.Add(new Tuple<int, string>(13, "255,0,0"));
            originalRGBs.Add(new Tuple<int, string>(14, "255,0,255"));
            originalRGBs.Add(new Tuple<int, string>(15, "0,255,255"));
            originalRGBs.Add(new Tuple<int, string>(16, "128,0,0"));
            originalRGBs.Add(new Tuple<int, string>(17, "0,128,0"));
            originalRGBs.Add(new Tuple<int, string>(18, "0,0,128"));
            originalRGBs.Add(new Tuple<int, string>(19, "128,128,0"));
            originalRGBs.Add(new Tuple<int, string>(20, "128,0,128"));
            originalRGBs.Add(new Tuple<int, string>(21, "0,128,128"));
            originalRGBs.Add(new Tuple<int, string>(22, "192,192,192"));
            originalRGBs.Add(new Tuple<int, string>(23, "128,128,128"));
            originalRGBs.Add(new Tuple<int, string>(24, "153,153,255"));
            originalRGBs.Add(new Tuple<int, string>(25, "153,51,102"));
            originalRGBs.Add(new Tuple<int, string>(26, "255,255,204"));
            originalRGBs.Add(new Tuple<int, string>(27, "204,255,255"));
            originalRGBs.Add(new Tuple<int, string>(28, "102,0,102"));
            originalRGBs.Add(new Tuple<int, string>(29, "255,128,128"));
            originalRGBs.Add(new Tuple<int, string>(30, "0,102,204"));
            originalRGBs.Add(new Tuple<int, string>(31, "204,204,255"));
            originalRGBs.Add(new Tuple<int, string>(32, "0,0,128"));
            originalRGBs.Add(new Tuple<int, string>(33, "255,0,255"));
            originalRGBs.Add(new Tuple<int, string>(34, "255,255,0"));
            originalRGBs.Add(new Tuple<int, string>(35, "0,255,255"));
            originalRGBs.Add(new Tuple<int, string>(36, "128,0,128"));
            originalRGBs.Add(new Tuple<int, string>(37, "128,0,0"));
            originalRGBs.Add(new Tuple<int, string>(38, "0,128,128"));
            originalRGBs.Add(new Tuple<int, string>(39, "0,0,255"));
            originalRGBs.Add(new Tuple<int, string>(40, "0,204,255"));
            originalRGBs.Add(new Tuple<int, string>(41, "204,255,255"));
            originalRGBs.Add(new Tuple<int, string>(42, "204,255,204"));
            originalRGBs.Add(new Tuple<int, string>(43, "255,255,153"));
            originalRGBs.Add(new Tuple<int, string>(44, "153,204,255"));
            originalRGBs.Add(new Tuple<int, string>(45, "255,153,204"));
            originalRGBs.Add(new Tuple<int, string>(46, "204,153,255"));
            originalRGBs.Add(new Tuple<int, string>(47, "255,204,153"));
            originalRGBs.Add(new Tuple<int, string>(48, "51,102,255"));
            originalRGBs.Add(new Tuple<int, string>(49, "51,204,204"));
            originalRGBs.Add(new Tuple<int, string>(50, "153,204,0"));
            originalRGBs.Add(new Tuple<int, string>(51, "255,204,0"));
            originalRGBs.Add(new Tuple<int, string>(52, "255,153,0"));
            originalRGBs.Add(new Tuple<int, string>(53, "255,102,0"));
            originalRGBs.Add(new Tuple<int, string>(54, "102,102,153"));
            originalRGBs.Add(new Tuple<int, string>(55, "150,150,150"));
            originalRGBs.Add(new Tuple<int, string>(56, "0,51,102"));
            originalRGBs.Add(new Tuple<int, string>(57, "51,153,102"));
            originalRGBs.Add(new Tuple<int, string>(58, "0,51,0"));
            originalRGBs.Add(new Tuple<int, string>(59, "51,51,0"));
            originalRGBs.Add(new Tuple<int, string>(60, "153,51,0"));
            originalRGBs.Add(new Tuple<int, string>(61, "153,51,102"));
            originalRGBs.Add(new Tuple<int, string>(62, "51,51,153"));
            originalRGBs.Add(new Tuple<int, string>(63, "51,51,51"));

        }

        #endregion

        #region 自定义颜色处理

        /// <summary>
        /// 低版本设置自定义颜色
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgbs"></param>
        private static void SetCustomColor(this HSSFWorkbook workBook, IEnumerable<string> rgbs)
        {
            SetOriginalRGB();
            // 获取调色板
            HSSFPalette pattern = workBook.GetCustomPalette();
            short indexed = 8;

            foreach (var rgb in rgbs)
            {
                if (indexed > 63)
                    indexed = 8;

                short tempIndex = pattern.SetCustomColor(rgb, indexed);
                if (tempIndex == -1)
                    continue;
                indexed = tempIndex;
                indexed++;
            }
        }

        private static void SetCustomColor(this HSSFWorkbook workBook, IEnumerable<ColorEntity> listColor)
        {
            SetOriginalRGB();
            // 获取调色板
            HSSFPalette pattern = workBook.GetCustomPalette();
            short indexed = 8;

            foreach (var color in listColor)
            {
                if (indexed > 63)
                    indexed = 8;

                short tempIndex = pattern.SetCustomColor(color.RGB, color.Index);
                if (tempIndex == -1)
                    continue;
                indexed = tempIndex;
                indexed++;
            }
        }

        /// <summary>
        /// 设置自定义颜色
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgb"></param>
        private static void SetCustomColor(this HSSFWorkbook workBook, string rgb)
        {
            SetOriginalRGB();

            HSSFPalette pattern = workBook.GetCustomPalette();
            pattern.SetCustomColor(rgb, -1);
        }

        /// <summary>
        /// 设置颜色
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="colorType"></param>
        private static void SetCustomColor(this HSSFWorkbook workBook, ColorType colorType)
        {
            string rgb = ExcelExtend.GetColor(colorType).Item2;
            workBook.SetCustomColor(rgb);
        }

        /// <summary>
        /// 设置颜色
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="rgb"></param>
        /// <param name="indexed"></param>
        /// <returns></returns>
        private static short SetCustomColor(this HSSFPalette pattern, string rgb, short indexed)
        {
            if (string.IsNullOrEmpty(rgb))
                return -1;
            string[] colors = rgb.Split(',');
            if (colors.Length != 3)
                return -1;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            // 处理RGB数据
            bool result = DealRGB(colors, ref red, ref green, ref blue);
            if (result == false)
                return -1;
            var temp = pattern.FindColor(red, green, blue);
            if (temp != null)
                return temp.Indexed;

            if (indexed == -1)
                indexed = 8;
            // 此位置下画板 原始rgb颜色
            string originalColor = originalRGBs.Where(t => t.Item1 == indexed).Select(t => t.Item2).FirstOrDefault();
            // 此位置下画板 rgb颜色
            string originalColor1 = string.Join(",", pattern.GetColor(indexed).RGB);
            // 如果两种颜色不一致，说明此位置已经设置了其他颜色，换个位置去设置
            if (originalColor != originalColor1)
            {
                indexed++;
                // 循环判断此位置颜色是否是原始颜色，如果是则设置，否则找其他位置
                // 如果此位置已经是最后位置了，则使用开始位置设置
                while (originalColor != originalColor1 || indexed < 64)
                {
                    originalColor = originalRGBs.Where(t => t.Item1 == indexed).Select(t => t.Item2).FirstOrDefault();
                    originalColor1 = string.Join(",", pattern.GetColor(indexed).RGB);
                    if (originalColor == originalColor1)
                        break;
                    indexed++;
                }
                if (indexed > 63)
                    indexed = 8;
            }

            pattern.SetColorAtIndex(indexed, red, green, blue);
            return indexed;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgbs"></param>
        private static void SetCustomColor(this XSSFWorkbook workBook, IEnumerable<string> rgbs)
        {
            // 获取调色板
            XSSFColor color = null;
            short indexed = 8;
            foreach (var rgb in rgbs)
            {
                if (indexed > 63)
                    indexed = 8;
                if (string.IsNullOrEmpty(rgb))
                    continue;
                string[] colors = rgb.Split(',');
                if (colors.Length != 3)
                    continue;
                byte red = 0;
                byte green = 0;
                byte blue = 0;
                // 处理RGB数据
                bool result = DealRGB(colors, ref red, ref green, ref blue);
                if (result == false)
                    continue;
                byte[] bytes = { red, green, blue };
                color = new XSSFColor();
                color.SetRgb(bytes);
                color.Indexed = indexed;
                indexed++;
            }
        }

        /// <summary>
        /// 获取自定义颜色位置
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgb"></param>
        /// <returns></returns>
        private static short GetCustomColor(this HSSFWorkbook workBook, string rgb)
        {
            SetOriginalRGB();
            short indexed = defaultColorIndexed;
            if (string.IsNullOrEmpty(rgb))
                return indexed;
            string[] colors = rgb.Split(',');
            if (colors.Length != 3)
                return indexed;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            bool result = DealRGB(colors, ref red, ref green, ref blue);
            if (result == false)
                return indexed;
            HSSFPalette pattern = workBook.GetCustomPalette();
            NPOI.HSSF.Util.HSSFColor hssfColor = pattern.FindColor(red, green, blue);
            if (hssfColor == null)
                return pattern.SetCustomColor(rgb, -1);
            indexed = hssfColor.Indexed;
            return indexed;
        }

        /// <summary>
        /// 高版本获取自定义颜色
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="rgb"></param>
        /// <returns></returns>
        private static short GetCustomColor(this XSSFWorkbook workBook, string rgb)
        {
            short indexed = defaultColorIndexed;
            if (string.IsNullOrEmpty(rgb))
                return indexed;
            string[] colors = rgb.Split(',');
            if (colors.Length != 3)
                return indexed;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            bool result = DealRGB(colors, ref red, ref green, ref blue);
            if (result == false)
                return indexed;
            byte[] bytes = { red, green, blue };
            XSSFColor color = new XSSFColor();
            color.SetRgb(bytes);
            indexed = color.Indexed;

            return indexed;
        }

        /// <summary>
        /// 处理RGB
        /// </summary>
        /// <param name="colors"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        private static bool DealRGB(string[] colors, ref byte red, ref byte green, ref byte blue)
        {
            bool result = true;
            red = 0;
            green = 0;
            blue = 0;
            if (byte.TryParse(colors[0], out red) &&
                byte.TryParse(colors[1], out green) &&
                byte.TryParse(colors[2], out blue))
            {
                // 如果超出255，则默认255；如果小于0，则默认0
                if (red > 255)
                    red = 255;
                if (red < 0)
                    red = 0;
                if (green > 255)
                    green = 255;
                if (green < 0)
                    green = 0;
                if (blue > 255)
                    blue = 255;
                if (blue < 0)
                    blue = 0;
            }
            else
                result = false;

            return result;
        }

        #endregion
    }
}
