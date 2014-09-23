using System.Collections.Generic;
using Ric.Util;

namespace Ric.Db
{
    public class TableManager
    {
        private const string TableListFile = "DbTables.xml";
        private static List<DbTable> _tableList;

        //To do: Get from database
        public static List<DbTable> GetTableList()
        {
            return _tableList ??
                   (_tableList = ConfigUtil.ReadConfig(TableListFile, typeof (List<DbTable>)) as List<DbTable>);
        }
    }

    public class DbTable
    {
        public string Name { get; set; }
        public string Market { get; set; }
        public string Description { get; set; }
    }
}
