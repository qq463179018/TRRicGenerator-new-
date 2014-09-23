using System;
using Ric.FileLib;
using Ric.FileLib.Attribute;

namespace Ric.FileLib.Entry
{
    /// <summary>
    /// Basic representation of an Idn bulk file entry
    /// </summary>
    public class IdnEntry : AEntry
    {
        [TitleName("SYMBOL")]
        public string Symbol { get; set; }

        [TitleName("DSPLY_NAME")]
        public string DisplayName { get; set; }

        [TitleName("OFFCL_CODE")]
        public string OfficialCode { get; set; }

        [TitleName("EX_SYMBOL")]
        public string ExSymbol { get; set; }

        [TitleName("BCKGRNDPAG")]
        public string BackgroundPage { get; set; }

        [TitleName("BCAST_REF")]
        public string Broadcast { get; set; }

        [TitleName("#INSTMOD_EXPIR_DATE")]
        public string TickerSymbol { get; set; }

        [TitleName("#INSTMOD_LONGLINK1")]
        public int RoundLotSize { get; set; }

        [TitleName("#INSTMOD_LONGLINK2")]
        public string BaseAsset { get; set; }
        
        [TitleName("#INSTMOD_MATUR_DATE")]
        public DateTime MatureDate { get; set; }
        
        [TitleName("#INSTMOD_OFFC_CODE2")]
        public string OfficialCode2 { get; set; }

        [TitleName("#INSTMOD_STRIKE_PRC")]
        public double StrikePrice { get; set; }

        [TitleName("#INSTMOD_WNT_RATIO")]
        public double WarrantRatio { get; set; }

        [TitleName("#INSTMOD_MNEMONIC")]
        public string Mnemonic { get; set; }

        [TitleName("#INSTMOD_TDN_SYMBOL")]
        public string TradingSymbol { get; set; }

        [TitleName("#INSTMOD_LONGLINK3")]
        public string Longlink3 { get; set; }

        [TitleName("#INSTMOD_GV1_DATE")]
        public DateTime GivenDate { get; set; }

        [TitleName("#INSTMOD_PUTCALLIND")]
        public string PutCall { get; set; }

        [TitleName("REF_COUNT")]
        public string RefCount { get; set; }

        [TitleName("EXL_NAME")]
        public string ExlName { get; set; }

        [TitleName("LINK_1")]
        public string Link1 { get; set; }

        [TitleName("LINK_2")]
        public string Link2 { get; set; }
    }
}
