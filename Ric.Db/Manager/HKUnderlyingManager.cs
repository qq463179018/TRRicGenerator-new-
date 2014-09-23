using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data;

namespace Ric.Db.Manager
{

    public class HKUnderlyingManager : ManagerBase
    {
        private const string ETI_UNDERLYING_TABLE_NAME = "ETI_HK_Underlying";

        public static HKUnderlyingInfo SelectUnderlyingInfoByUnderlying(string underlying)
        {
            try
            {
                if ((underlying + string.Empty).Trim().Length == 0)
                    return null;

                string where = string.Format("where Underlying = '{0}'", underlying);
                DataTable dt = ManagerBase.Select(ETI_UNDERLYING_TABLE_NAME, new string[] { "*" }, where);

                if (dt == null || dt.Rows.Count <= 0)
                    return null;

                DataRow dr = dt.Rows[0];
                HKUnderlyingInfo hkUnderlying = new HKUnderlyingInfo();
                hkUnderlying.ID = Convert.ToString(dr["ID"]);
                hkUnderlying.Underlying = Convert.ToString(dr["Underlying"]);
                hkUnderlying.BCAST_REF = Convert.ToString(dr["BCAST_REF"]);
                hkUnderlying.INSTMOD_GN_TX20_6 = Convert.ToString(dr["INSTMOD_GN_TX20_6"]);
                hkUnderlying.INSTMOD_GN_TX20_7 = Convert.ToString(dr["INSTMOD_GN_TX20_7"]);
                hkUnderlying.INSTMOD_GN_TX20_12 = Convert.ToString(dr["INSTMOD_GN_TX20_12"]);
                hkUnderlying.INSTMOD_LONGLINK2 = Convert.ToString(dr["INSTMOD_LONGLINK2"]);
                hkUnderlying.INSTMOD_LONGLINK6 = Convert.ToString(dr["INSTMOD_LONGLINK6"]);
                hkUnderlying.INSTMOD_UNDERLYING = Convert.ToString(dr["INSTMOD_UNDERLYING"]);

                return hkUnderlying;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
