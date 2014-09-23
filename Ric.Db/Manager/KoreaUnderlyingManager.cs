using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data;

namespace Ric.Db.Manager
{
    public class KoreaUnderlyingManager : ManagerBase
    {
        const string ETI_KOREA_UNDERLYING_TABLE_NAME = "ETI_Korea_Underlying";

        /// <summary>
        /// Select record of underlying in DB with korea name.
        /// </summary>
        /// <param name="koreaName"></param>
        /// <returns></returns>
        public static KoreaUnderlyingInfo SelectUnderlying(string koreaName)
        {
            return SelectUnderlying(koreaName, KoreaNameType.KoreaName);
        }

        /// <summary>
        /// Select record of underlying in DB with korea name and name type(FM1, FM2 or Drop)
        /// </summary>
        /// <param name="koreaName"></param>
        /// <returns></returns>
        public static KoreaUnderlyingInfo SelectUnderlying(string koreaName, KoreaNameType type)
        {
            string condition = "";
            if (type.Equals(KoreaNameType.KoreaName))
            {
                condition = "where KoreaName =N'" + koreaName + "'";
            }
            else if (type.Equals(KoreaNameType.KoreaNameForFM2))
            {
                condition = "where KoreaNameFM2 =N'" + koreaName + "'or KoreaNameDrop =N'" + koreaName + "' or UnderlyingName =N'" + koreaName + "'";
            }
            else if (type.Equals(KoreaNameType.KoreaNameForDrop))
            {
                condition = "where KoreaNameFM2 =N'" + koreaName + "'or KoreaNameDrop =N'" + koreaName + "' or UnderlyingName =N'" + koreaName + "'";         
            }
            else
            {
                condition = "where KoreaNameFM2 =N'" + koreaName + "'or KoreaNameDrop =N'" + koreaName + "' or UnderlyingName =N'" + koreaName + "'";         
            }

            DataTable dt = ManagerBase.Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            KoreaUnderlyingInfo underlying = new KoreaUnderlyingInfo();
            DataRow dr = dt.Rows[0];
            underlying.BNDUnderlying = Convert.ToString(dr["BNDUnderlying"]);
            underlying.BodyGroupCommonName = Convert.ToString(dr["BodyGroupCommonName"]);
            underlying.IDNDisplayNamePart = Convert.ToString(dr["IDNDisplayNamePart"]);
            underlying.KoreaName = Convert.ToString(dr["KoreaName"]);
            underlying.KoreaNameDrop = Convert.ToString(dr["KoreaNameDrop"]);
            underlying.KoreaNameFM2 = Convert.ToString(dr["KoreaNameFM2"]);
            underlying.NDATCUnderlyingTitle = Convert.ToString(dr["NDATCUnderlyingTitle"]);
            underlying.QACommonNamePart = Convert.ToString(dr["QACommonNamePart"]);
            underlying.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
            underlying.UnderlyingRIC = Convert.ToString(dr["UnderlyingRIC"]);
            underlying.ISIN = Convert.ToString(dr["ISIN"]);
            return underlying;
        }

        public static KoreaUnderlyingInfo SelectUnderlyingByISIN(string isin)
        {
            string condition = "where ISIN ='" + isin + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            KoreaUnderlyingInfo underlying = new KoreaUnderlyingInfo();
            DataRow dr = dt.Rows[0];
            underlying.BNDUnderlying = Convert.ToString(dr["BNDUnderlying"]);
            underlying.BodyGroupCommonName = Convert.ToString(dr["BodyGroupCommonName"]);
            underlying.IDNDisplayNamePart = Convert.ToString(dr["IDNDisplayNamePart"]);
            underlying.KoreaName = Convert.ToString(dr["KoreaName"]);
            underlying.KoreaNameDrop = Convert.ToString(dr["KoreaNameDrop"]);
            underlying.KoreaNameFM2 = Convert.ToString(dr["KoreaNameFM2"]);
            underlying.NDATCUnderlyingTitle = Convert.ToString(dr["NDATCUnderlyingTitle"]);
            underlying.QACommonNamePart = Convert.ToString(dr["QACommonNamePart"]);
            underlying.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
            underlying.UnderlyingRIC = Convert.ToString(dr["UnderlyingRIC"]);            
            return underlying;        
        }


        public static void UpdateKoreaNameFM2(string koreaNameFM2, string isin)
        {
            DataTable dt = Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, "where isin = '" + isin + "'");
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }
            
            foreach (DataRow row in dt.Rows)
            { 
                row["KoreaNameFM2"] = koreaNameFM2;               
            }
           
            UpdateDbTable(dt, ETI_KOREA_UNDERLYING_TABLE_NAME);  
        }

        public static void UpdateKoreaNameDrop(string koreaNameDrop, string isin)
        {
            DataTable dt = Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, "where isin = '" + isin + "'");
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }

            foreach (DataRow row in dt.Rows)
            {
                row["KoreaNameDrop"] = koreaNameDrop;
            }

            UpdateDbTable(dt, ETI_KOREA_UNDERLYING_TABLE_NAME);  
        
        }


        /// <summary>
        /// Check if the table contains a same display name for new underlying.
        /// </summary>
        /// <param name="displayName">display name</param>
        /// <returns>true or false</returns>
        public static bool ExsitDisplayName(string displayName)
        {
            string condition = "where IDNDisplayNamePart = '" + displayName + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static void UpdateUnderlying(KoreaUnderlyingInfo underlying)
        {
            DataTable dt = Select(ETI_KOREA_UNDERLYING_TABLE_NAME, new string[] { "*" }, "where KoreaName = N'" + underlying.KoreaName + "'");
            DataRow dr = null;
            if (dt == null || dt.Rows.Count == 0)
            {
                dr = dt.NewRow();
                dr["BNDUnderlying"] = underlying.BNDUnderlying;
                dr["BodyGroupCommonName"] = underlying.BodyGroupCommonName;
                dr["IDNDisplayNamePart"] = underlying.IDNDisplayNamePart;
                dr["KoreaName"] = underlying.KoreaName;
                dr["KoreaNameDrop"] = underlying.KoreaNameDrop;
                dr["KoreaNameFM2"] = underlying.KoreaNameFM2;
                dr["NDATCUnderlyingTitle"] = underlying.NDATCUnderlyingTitle;
                dr["QACommonNamePart"] = underlying.QACommonNamePart;
                dr["UnderlyingName"] = underlying.UnderlyingName;
                dr["UnderlyingRIC"] = underlying.UnderlyingRIC;
                dr["ISIN"] = underlying.ISIN;
                dt.Rows.Add(dr);
            }
            else
            {
                foreach(DataRow row in dt.Rows)
                {
                    row["BNDUnderlying"] = underlying.BNDUnderlying;
                    row["BodyGroupCommonName"] = underlying.BodyGroupCommonName;
                    row["IDNDisplayNamePart"] = underlying.IDNDisplayNamePart;
                    row["KoreaName"] = underlying.KoreaName;
                    row["KoreaNameDrop"] = underlying.KoreaNameDrop;
                    row["KoreaNameFM2"] = underlying.KoreaNameFM2;
                    row["NDATCUnderlyingTitle"] = underlying.NDATCUnderlyingTitle;
                    row["QACommonNamePart"] = underlying.QACommonNamePart;
                    row["UnderlyingName"] = underlying.UnderlyingName;
                    row["UnderlyingRIC"] = underlying.UnderlyingRIC;
                    row["ISIN"] = underlying.ISIN;
                }
            }
            UpdateDbTable(dt, ETI_KOREA_UNDERLYING_TABLE_NAME);  
        }
    }

    public enum KoreaNameType
    {
        KoreaName,
        KoreaNameForFM2,
        KoreaNameForDrop,
        KoreaUnderlyingName
    }
}
