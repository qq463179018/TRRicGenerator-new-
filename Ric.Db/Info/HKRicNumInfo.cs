using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class HKRicNumInfo
    {
        public string LauchDate { get; set; }
        public int CBBCRicNum { get; set; }
        public int WarrantRicNum { get; set; }

        public HKRicNumInfo(string date, int CBBCRicNum, int WarrantRicNum)
        {
            this.LauchDate = date;
            this.CBBCRicNum = CBBCRicNum;
            this.WarrantRicNum = WarrantRicNum;
        }

        public HKRicNumInfo()
        { }
    }
}
