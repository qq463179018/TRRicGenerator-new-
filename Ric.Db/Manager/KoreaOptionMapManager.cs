using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data;

namespace Ric.Db.Manager
{
    public class KoreaOptionMapManager : ManagerBase
    {
        const string ETI_KOREA_CODEMAP_TABLE_NAME = "ETI_Korea_CodeMap";
        const string ETI_KOREA_OPTION_LTD_TABLE_NAME = "ETI_Korea_OptionLTD";

        //  public static Hashset 

        public static List<KoreaCodeMapInfo> SelectCodeMapByType(KoreaCodeMapType codeType)
        {
            string condition = string.Format(" where Type = '{0}'", (int)codeType);
            DataTable dt = Select(ETI_KOREA_CODEMAP_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            List<KoreaCodeMapInfo> codeMap = new List<KoreaCodeMapInfo>();
            foreach (DataRow dr in dt.Rows)
            {
                KoreaCodeMapInfo codeItem = new KoreaCodeMapInfo();
                codeItem.Name = Convert.ToString(dr["Name"]).Trim();
                codeItem.Code = Convert.ToString(dr["Code"]).Trim();
                codeItem.Type = codeType;
                codeMap.Add(codeItem);
            }
            return codeMap;
        }

        public static KoreaCodeMapInfo SelectOneYearCode(int year)
        {
            string condition = string.Format(" where Type = {0} and Name = '{1}'", (int)KoreaCodeMapType.Year, year );
            DataTable dt = Select(ETI_KOREA_CODEMAP_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            KoreaCodeMapInfo codeMap = new KoreaCodeMapInfo();
            DataRow dr = dt.Rows[0];
            codeMap.Name = Convert.ToString(dr["Name"]).Trim();
            codeMap.Code = Convert.ToString(dr["Code"]).Trim();
            return codeMap;
        }

        public static List<KoreaOptionLastTradingDayInfo> SelectLastTradingDayByYear(string year)
        {
            string condition = string.Format(" where Year = '{0}' order by LastTradingDay ", year);
            DataTable dt = Select(ETI_KOREA_OPTION_LTD_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            List<KoreaOptionLastTradingDayInfo> lastTradingDayMap = new List<KoreaOptionLastTradingDayInfo>();
            foreach (DataRow dr in dt.Rows)
            {
                KoreaOptionLastTradingDayInfo lastTradingDayItem = new KoreaOptionLastTradingDayInfo();
                lastTradingDayItem.Year = year;
                lastTradingDayItem.Month = Convert.ToString(dr["Month"]).Trim();
                lastTradingDayItem.LastTradingDay = Convert.ToDateTime(dr["LastTradingDay"]);
                lastTradingDayItem.LastTradingDayForUSDOPT = Convert.ToDateTime(dr["LastTradingDayForUSDOPT"]);
                lastTradingDayItem.ContractMonthNumber = Convert.ToString(dr["ContractMonthNumber"]).Trim();
                lastTradingDayMap.Add(lastTradingDayItem);
            }
            return lastTradingDayMap;
        }

        public static bool CheckLastTradingDay(string currentDate)
        {
            string condition = string.Format("where LastTradingDay = '{0}'", currentDate);
            DataTable dt = Select(ETI_KOREA_OPTION_LTD_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static bool CheckLaterThanLastTradingDay(DateTime currentDate)
        {
            string year = currentDate.Year.ToString();
            string month = ((StockOptionMonth)currentDate.Month).ToString();
            string today = currentDate.ToString("yyyy-MM-dd");
            string condition = string.Format("where LastTradingDay <= '{0}' and Year = '{1}' and Month = '{2}'", today , year, month);
            DataTable dt = Select(ETI_KOREA_OPTION_LTD_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static bool CheckLaterThanLastTradingDayUSD(DateTime currentDate)
        {
            string year = currentDate.Year.ToString();
            string month = ((StockOptionMonth)currentDate.Month).ToString();
            string today = currentDate.ToString("yyyy-MM-dd");
            string condition = string.Format("where LastTradingDayForUSDOPT <= '{0}' and Year = '{1}' and Month = '{2}'", today, year, month);
            DataTable dt = Select(ETI_KOREA_OPTION_LTD_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static KoreaOptionLastTradingDayInfo SelectLastTradingDayByYearMonth(string year, string month)
        {
            string condition = string.Format(" where Year = '{0}' and Month = '{1}' ", year, month);
            DataTable dt = Select(ETI_KOREA_OPTION_LTD_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            DataRow dr = dt.Rows[0];

            KoreaOptionLastTradingDayInfo lastTradingDayItem = new KoreaOptionLastTradingDayInfo();
            lastTradingDayItem.Year = year;
            lastTradingDayItem.Month = month;
            lastTradingDayItem.LastTradingDay = Convert.ToDateTime(dr["LastTradingDay"]);
            lastTradingDayItem.LastTradingDayForUSDOPT = Convert.ToDateTime(dr["LastTradingDayForUSDOPT"]);
            lastTradingDayItem.ContractMonthNumber = Convert.ToString(dr["ContractMonthNumber"]).Trim();

            return lastTradingDayItem;

        }
    }
}
