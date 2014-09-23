using System;
using Ric.FileLib.Attribute;

namespace Ric.FileLib.Entry
{
    /// <summary>
    /// Basic representation of a Nda bulk file entry
    /// </summary>
    public class NdaEntry : AEntry
    {
        [TitleName("TYPE")]
        public string Type { get; set; }

        [TitleName("CATEGORY")]
        public string Category { get; set; }

        [TitleName("EXCHANGE")]
        public string Exchange { get; set; }

        [TitleName("CURRENCY")]
        public string Currency { get; set; }

        [TitleName("ASSET COMMON NAME")]
        public string AssetCommonName { get; set; }

        [TitleName("ASSET SHORT NAME")]
        public string AssetShortName { get; set; }

        [TitleName("EXPIRY DATE")]
        public DateTime ExpiryDate { get; set; }

        [TitleName("STRIKE PRICE")]
        public int StrikePrice { get; set; }

        [TitleName("TICKER SYMBOL")]
        public string TickerSymbol { get; set; }

        [TitleName("TRADING SEGMENT")]
        public string TradingSegment { get; set; }

        [TitleName("TAG")]
        public string Tag { get; set; }

        [TitleName("ROUND LOT SIZE")]
        public int RoundLotSize { get; set; }

        [TitleName("BASE ASSET")]
        public string BaseAsset { get; set; }

        [TitleName("PRIMARY TRADABLE MARKET QUOTE")]
        public string PrimaryTradableMarketQuote { get; set; }

        [TitleName("CALL PUT OPTION")]
        public string CallPutOption { get; set; }

        [TitleName("EQUITY FIRST TRADING DAY")]
        public DateTime EquityFirstTradingDay { get; set; }

        [TitleName("SETTLEMENT PERIOD")]
        public string SettlementPeriod { get; set; }
    }
}
