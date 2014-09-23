using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class KoreaCodeMapInfo
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public KoreaCodeMapType Type { get; set; }
    }

    public enum KoreaCodeMapType : int
    {
        Year = 0,
        Call = 1,
        Put = 2
    }

    public class KoreaOptionLastTradingDayInfo
    {
        public string Year { get; set; }
        public string Month { get; set; }
        public DateTime LastTradingDay { get; set; }
        public DateTime LastTradingDayForUSDOPT { get; set; }
        public string ContractMonthNumber { get; set; }    
    }

    public enum StockOptionMonth
    {
        NoneOfUse, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OCT, NOV, DEC
    }
}
