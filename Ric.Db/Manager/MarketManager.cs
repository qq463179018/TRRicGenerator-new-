using System;
using System.Collections.Generic;
using System.Linq;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class MarketManager : ManagerBase
    {
        public static List<Market> GetAllMarkets(EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from market in ctx.Markets
                        select market).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
