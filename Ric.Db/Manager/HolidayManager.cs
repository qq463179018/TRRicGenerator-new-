using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Ric.Db.Manager
{

    public class HolidayInfo
    {
        public int MarketId { get; set; }

        public string HolidayDate { get; set; }

        public string Comment { get; set; }
    }

    public class HolidayManager : ManagerBase
    {
        private const string ETI_HOLIDAY_TABLE_NAME = "ETI_Holiday";

        public static int UpdateHoliday(HolidayInfo holiday)
        {
            string where = string.Format("where HolidayDate = '{0}' and MarketId = {1}", holiday.HolidayDate, holiday.MarketId);
            DataTable dt = Select(ETI_HOLIDAY_TABLE_NAME, new[] { "*" }, where);
            if (dt == null || dt.Rows.Count > 0)
            {
                return 0;
            }
            DataRow dr = dt.NewRow();
            dr["HolidayDate"] = holiday.HolidayDate;
            dr["MarketId"] = holiday.MarketId;
            dr["Comment"] = holiday.Comment;
            dt.Rows.Add(dr);
            return UpdateDbTable(dt, ETI_HOLIDAY_TABLE_NAME);
        }

        public static int UpdateHoliday(List<HolidayInfo> holiday)
        {
            return holiday.Sum(item => UpdateHoliday(item));
        }

        public static List<DateTime> SelectHoliday(int marketId)
        { 
            return SelectHoliday(DateTime.Today.Year, marketId);
        }

        public static List<DateTime> SelectHoliday(int year, int marketId)
        {
            List<DateTime> holidayList = new List<DateTime>();
            string where = string.Format("where DATEPART(yyyy,HolidayDate) = '{0}' and MarketId = {1}", year, marketId);

            DataTable dt = ManagerBase.Select(ETI_HOLIDAY_TABLE_NAME, new[] { "*" }, where);
            if (dt == null || dt.Rows.Count == 0)
            {
                return holidayList;
            }
            holidayList.AddRange(from DataRow dr in dt.Rows select Convert.ToDateTime(dr["HolidayDate"]));
            return holidayList;
        }
       
    }
}
