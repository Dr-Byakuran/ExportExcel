
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
    /// 字体类型
    /// </summary>
    public enum FontNameType
    {
        Arial_Unicode_MS = 1,

        Malgun_Gothic = 2,

        Malgun_Gothic_Semilight = 3,

        Meiryo = 4,

        Meiryo_UI = 5,

        Microsoft_JhengHei = 6,

        Microsoft_JhengHei_Light = 7,

        Microsoft_JhengHei_UI = 8,

        Microsoft_JhengHei_UI_Light = 9,

        Microsoft_YaHei_UI = 10,

        Microsoft_YaHei_UI_Light = 11,

        MingLiU_HKSCS_ExtB = 12,

        MingLiU_ExtB = 13,

        MS_Gothic = 14,

        MS_PGothic = 15,

        MS_UI_Gothic = 16,

        PMingLiU_ExtB = 17,

        SimSun_ExtB = 18,

        Yu_Gothic = 19,

        Yu_Gothic_Light = 20,

        Yu_Gothic_Medium = 21,

        Yu_Gothic_UI = 22,

        Yu_Gothic_UI_Light = 23,

        Yu_Gothic_UI_Semibold = 24,

        Yu_Gothic_UI_Semilight = 25,

        等线 = 26,

        等线_Light = 27,

        方正兰亭超细黑简体 = 28,

        方正舒体 = 29,

        方正姚体 = 30,

        仿宋 = 31,

        黑体 = 32,

        华文彩云 = 33,

        华文仿宋 = 34,

        华文行楷 = 35,

        华文琥珀 = 36,

        华文楷体 = 37,

        华文隶书 = 38,

        华文宋体 = 39,

        华文细黑 = 40,

        华文新魏 = 41,

        华文中宋 = 42,

        楷体 = 43,

        隶书 = 44,

        宋体 = 45,

        微软雅黑 = 46,

        微软雅黑_Light = 47,

        新宋体 = 48,

        幼圆 = 49,

        Agency_FB = 50,

        Algerian = 51,

        Arial = 52,

        Arial_Black = 53,

        Arial_Narrow = 54,

        Arial_Rounded_MT_Bold = 55,

        Baskerville_Old_Face = 56,

        Bauhaus_93 = 57,

        Bell_MT = 58,

        Berlin_Sans_FB = 59,

        Berlin_Sans_FB_Demi = 60,

        Bernard_MT_Condensed = 61,

        Blackadder_ITC = 62,

        Bodoni_MT = 63,

        Bodoni_MT_Black = 64,

        Bodoni_MT_Condensed = 65,

        Bodoni_MT_Poster_Compressed = 66,

        Book_Antiqua = 67,

        Bookman_Old_Style = 68,

        Bookshelf_Symbol_7 = 69,

        Bradley_Hand_ITC = 70,

        Britannic_Bold = 71,

        Broadway = 72,

        Brush_Script_MT = 73,

        Buxton_Sketch = 74,

        Calibri = 75,

        Calibri_Light = 76,

        Californian_FB = 77,

        Calisto_MT = 78,

        Cambria = 79,

        Cambria_Math = 80,

        Candara = 81,

        Castellar = 82,

        Centaur = 83,

        Century = 74,

        Century_Gothic = 75,

        Century_Schoolbook = 76,

        Chiller = 77,

        Colonna_MT = 78,

        Comic_Sans_MS = 79,

        Consolas = 80,

        Constantia = 81,

        Cooper_Black = 82,

        Copperplate_Gothic_Bold = 83,

        Copperplate_Gothic_Light = 84,

        Corbel = 85,

        Courier_New = 86,

        Curlz_MT = 87,

        Ebrima = 88,

        Edwardian_Script_ITC = 89,

        Elephant = 90,

        Engravers_MT = 91,

        Eras_Bold_ITC = 92,

        Eras_Demi_ITC = 93,

        Eras_Light_ITC = 94,

        Eras_Medium_ITC = 95,

        Felix_Titling = 96,

        Footlight_MT_Light = 97,

        Forte = 98,

        Franklin_Gothic_Book = 99,

        Franklin_Gothic_Demi = 100,

        Franklin_Gothic_Demi_Cond = 101,

        Franklin_Gothic_Heavy = 102,

        Franklin_Gothic_Medium = 103,

        Franklin_Gothic_Medium_Cond = 104,

        Freestyle_Script = 105,

        French_Script_MT = 106,

        Gabriola = 107,

        Gadugi = 108,

        Garamond = 109,

        Georgia = 110,

        Gigi = 111,

        Gill_Sans_MT = 112,

        Gill_Sans_MT_Condensed = 113,

        Gill_Sans_MT_Ext_Condensed_Bold = 114,

        Gill_Sans_Ultra_Bold = 115,

        Gill_Sans_Ultra_Bold_Condensed = 116,

        Gloucester_MT_Extra_Condensed = 117,

        Goudy_Old_Style = 118,

        Goudy_Stout = 119,

        Haettenschweiler = 120,

        Harlow_Solid_Italic = 121,

        Harrington = 122,

        High_Tower_Text = 123,

        HoloLens_MDL2_Assets = 124,

        Impact = 125,

        Imprint_MT_Shadow = 126,

        Informal_Roman = 127,

        Javanese_Text = 128,

        Jokerman = 129,

        Juice_ITC = 130,

        Kristen_ITC = 131,

        Kunstler_Script = 132,

        Leelawadee_UI = 133,

        Leelawadee_UI_Semilight = 134,

        Lucida_Bright = 135,

        Lucida_Calligraphy = 136,

        Lucida_Console = 137,

        Lucida_Fax = 138,

        Lucida_Handwriting = 139,

        Lucida_Sans = 140,

        Lucida_Sans_Typewriter = 141,

        Lucida_Sans_Unicode = 142,

        Magneto = 143,

        Maiandra_GD = 144,

        Marlett = 145,

        Matura_MT_Script_Capitals = 146,

        Microsoft_Himalaya = 147,

        Microsoft_New_Tai_Lue = 148,

        Microsoft_PhagsPa = 149,

        Microsoft_Sans_Serif = 150,

        Microsoft_Tai_Le = 151,

        Microsoft_Yi_Baiti = 152,

        Mistral = 153,

        Modern_No20 = 154,

        Mongolian_Baiti = 155,

        Monotype_Corsiva = 156,

        MS_Outlook = 157,

        MS_Reference_Sans_Serif = 158,

        MS_Reference_Specialty = 159,

        MT_Extra = 160,

        MV_Boli = 161,

        Myanmar_Text = 162,

        Niagara_Engraved = 163,

        Niagara_Solid = 164,

        Nirmala_UI = 165,

        Nirmala_UI_Semilight = 166,

        OCR_A_Extended = 167,

        Old_English_Text_MT = 168,

        Onyx = 169,

        Palace_Script_MT = 170,

        Palatino_Linotype = 171,

        Papyrus = 172,

        Parchment = 173,

        Perpetua = 174,

        Perpetua_Titling_MT = 175,

        Playbill = 176,

        Poor_Richard = 177,

        Pristina = 178,

        Rage_Italic = 179,

        Ravie = 180,

        Rockwell = 181,

        Rockwell_Condensed = 182,

        Rockwell_Extra_Bold = 183,

        Script_MT_Bold = 184,

        Segoe_Marker = 195,

        Segoe_MDL2_Assets = 196,

        Segoe_Print = 197,

        Segoe_Script = 198,

        Segoe_UI = 199,

        Segoe_UI_Black = 200,

        Segoe_UI_Emoji = 201,

        Segoe_UI_Historic = 202,

        Segoe_UI_Light = 203,

        Segoe_UI_Semibold = 204,

        Segoe_UI_Symbol = 205,

        Showcard_Gothic = 206,

        Sitka_Banner = 207,

        Sitka_Display = 208,

        Sitka_Heading = 209,

        Sitka_Small = 210,

        Sitka_Subheading = 211,

        Sitka_Text = 212,

        SketchFlow_Print = 213,

        Snap_ITC = 214,

        Stencil = 215,

        Sylfaen = 216,

        Symbol = 217,

        Tahoma = 218,

        TeamViewer13 = 219,

        Tempus_Sans_ITC = 220,

        Times_New_Roman = 221,

        Trebuchet_MS = 222,

        Tw_Cen_MT = 223,

        Tw_Cen_MT_Condensed = 224,

        Tw_Cen_MT_Condensed_Extra_Bold = 225,

        Verdana = 226,

        Viner_Hand_ITC = 227,

        Vivaldi = 228,

        Vladimir_Script = 229,

        Webdings = 230,

        Wide_Latin = 231,

        Wingdings = 232,

        Wingdings2 = 233,

        Wingdings3 = 234,
    }
}
