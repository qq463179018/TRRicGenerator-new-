using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class KoreaRicNumInfo
    {
        public string LauchDate { get; set; }
        public int WarrantAddRicNum { get; set; }
        public int WarrantDropRicNum { get; set; }

        public KoreaRicNumInfo(string date, int WarrantAddRicNum, int WarrantDropRicNum)
        {
            this.LauchDate = date;
            this.WarrantAddRicNum = WarrantAddRicNum;
            this.WarrantDropRicNum = WarrantDropRicNum;
        }

        public KoreaRicNumInfo()
        { }
    }
}
