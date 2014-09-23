using Ric.FileLib;
using Ric.FileLib.Attribute;

namespace Ric.FileLib.Entry
{
    /// <summary>
    /// Basic representation of a Nda bulk file entry
    /// </summary>
    public class FmEntry : AEntry
    {
        [TitleName("ASSET COMMON NAME")]
        public string AssetCommonName { get; set; }

        [TitleName("ASSET SHORT NAME")]
        public string AssetShortName { get; set; }

        [TitleName("CURRENCY")]
        public string Currency { get; set; }

        [TitleName("EXCHANGE")]
        public string Exchange { get; set; }

        [TitleName("TYPE")]
        public string Type { get; set; }

        [TitleName("CATEGORY")]
        public string Category { get; set; }

        [TitleName("TICKER SYMBOL")]
        public string TickerSymbol { get; set; }

        [TitleName("ROUND LOT SIZE")]
        public int RoundLotSize { get; set; }

        [TitleName("BASE ASSET")]
        public string BaseAsset { get; set; }

        [TitleName("TAG")]
        public string Tag { get; set; }
    }
}
