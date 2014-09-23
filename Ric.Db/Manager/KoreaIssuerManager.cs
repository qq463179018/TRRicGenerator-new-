using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data;

namespace Ric.Db.Manager
{
    public class KoreaIssuerManager : ManagerBase
    {
        const string ETI_KOREA_ISSUER_TABLE_NAME = "ETI_Korea_Issuer";
                
        /// <summary>
        /// Select a record of Korea Issuer with Korea issuer name.
        /// </summary>
        /// <param name="koreaIssuerName">Korea issuer name</param>
        /// <returns>record of issuer</returns>
        public static KoreaIssuerInfo SelectIssuer(string koreaIssuerName)
        {
            string condition = "where KoreaIssuerName =N'" + koreaIssuerName + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_ISSUER_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            KoreaIssuerInfo issuer = new KoreaIssuerInfo();
            DataRow dr = dt.Rows[0];
            issuer.KoreaIssuerName = Convert.ToString(dr["KoreaIssuerName"]);
            issuer.BodyGroupCommonName =Convert.ToString( dr["BodyGroupCommonName"]);
            issuer.IssuerCode2 = Convert.ToString(dr["IssuerCode2"]);
            issuer.IssuerCode4 = Convert.ToString(dr["IssuerCode4"]);
            issuer.IssuerCompanyName = Convert.ToString(dr["IssuerCompanyName"]);
            issuer.IssuerName5 = Convert.ToString(dr["IssuerName5"]);
            issuer.NDAIssuerOrgid = Convert.ToString(dr["NDAIssuerOrgid"]);
            issuer.NDATCIssuerTitle = Convert.ToString(dr["NDATCIssuerTitle"]);
            
            return issuer;            
        }

        /// <summary>
        /// Select a record of Korea Issuer with Korea issuer code2
        /// </summary>
        /// <param name="issuerCode2">Korea oissuer code2</param>
        /// <returns></returns>
        public static KoreaIssuerInfo SelectIssuerByIssuerCode2(string issuerCode2)
        {
            string condition = "where IssuerCode2 =N'" + issuerCode2 + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_ISSUER_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            KoreaIssuerInfo issuer = new KoreaIssuerInfo();
            DataRow dr = dt.Rows[0];
            issuer.KoreaIssuerName = Convert.ToString(dr["KoreaIssuerName"]);
            issuer.BodyGroupCommonName = Convert.ToString(dr["BodyGroupCommonName"]);
            issuer.IssuerCode2 = Convert.ToString(dr["IssuerCode2"]);
            issuer.IssuerCode4 = Convert.ToString(dr["IssuerCode4"]);
            issuer.IssuerCompanyName = Convert.ToString(dr["IssuerCompanyName"]);
            issuer.IssuerName5 = Convert.ToString(dr["IssuerName5"]);
            issuer.NDAIssuerOrgid = Convert.ToString(dr["NDAIssuerOrgid"]);
            issuer.NDATCIssuerTitle = Convert.ToString(dr["NDATCIssuerTitle"]);

            return issuer;
        }

    }
}
