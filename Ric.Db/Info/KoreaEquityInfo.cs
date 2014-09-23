using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class KoreaEquityInfo
    {
        public string UpdateDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string Type { get; set; }
        public string RecordType { get; set; }
        public string FM { get; set; }
        public string IDNDisplayName { get; set; }        
        public string ISIN { get; set; }
        public string Ticker { get; set; }
        public string BcastRef { get; set; }
        public string LegalName { get; set; }
        public string KoreaName { get; set; }
        public string Lotsize { get; set; } 
        public string Market { get; set; }
        public string Status { get; set; }

        public bool ExistsFM1 { get; set; }
        public bool IsGlobalETF { get; set; }

        //for exceptions and errors
        public string AnnouncementTime { get; set; }

        //for NDA file columns
        public string Category { get; set; }
        public string Exchange { get; set; }

        //for FM change columns
        public List<KoreaAddFMColumn> ChangeItems { get; set; }

        //for Name change
        public string OldIDNDisplayName { get; set; }
        public string OldLegalName { get; set; }
        public string OldKoreaName { get; set; }
        public bool IsRevised { get; set; }

        //for PRF ending 
        public string PrfEnd { get; set; }

        public KoreaEquityInfo()
        {
            UpdateDate = DateTime.Today.ToString("yyyy-MM-dd");
            EffectiveDate = DateTime.Today.ToString("yyyy-MM-dd");      
            Type = string.Empty;
            RecordType = string.Empty;
            FM = string.Empty;
            ISIN = string.Empty;
            Ticker = string.Empty;
            BcastRef = string.Empty;
            LegalName = string.Empty;
            KoreaName = string.Empty;              
            Lotsize = string.Empty;
            Status = string.Empty;
            ExistsFM1 = false;
            IsGlobalETF = false; 
            ChangeItems = new List<KoreaAddFMColumn>();
            IsRevised = false;
        }
    }

    public enum KoreaAddFMColumn : int
    {
        UpadatDate = 1,
        EffectiveDate = 2,
        RIC = 3,
        Type = 4,
        RecordType = 5,
        FM = 6,
        IDNDisplayName = 7,
        ISIN = 8,
        Ticker = 9,
        BCASTRef = 10,
        LegalName = 11,
        KoreaName = 12,
        LotSize = 13,
    }

 
}
