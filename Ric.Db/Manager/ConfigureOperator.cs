using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Manager
{
    public class ConfigureOperator
    {
        public static string GetEmaFileSaveDir()
        {
            try
            {
                var tableRows = new string[] { "ConfigValue" };
                var table = ManagerBase.Select("ConfigTable", tableRows, "where MarketName='Korea' AND ConfigType='EmaSaveFileDir'");
                string path = table.Rows[0][0].ToString();
                return path;
            }
            catch
            {
                return null;
            }
        }

        public static string GetGedaFileSaveDir()
        {
            try
            {
                var tableRows = new string[] { "ConfigValue" };
                var table = ManagerBase.Select("ConfigTable", tableRows, "where MarketName='Korea' AND ConfigType='GedaSaveFileDir'");
                string path = table.Rows[0][0].ToString();
                return path;
            }
            catch
            {
                return null;
            }
        }

        public static string GetNdaFileSaveDir()
        {
            try
            {
                var tableRows = new string[] { "ConfigValue" };
                var table = ManagerBase.Select("ConfigTable", tableRows, "where MarketName='Korea' AND ConfigType='NdaSaveFileDir'");
                string path = table.Rows[0][0].ToString();
                return path;
            }
            catch
            {
                return null;
            }
        }

        public static string GetGatsServer()
        {
            try
            {
                var tableRows = new string[] { "ConfigValue" };
                var table = ManagerBase.Select("ConfigTable", tableRows, "where MarketName='Other' AND ConfigType='GatsServer'");
                string ip = table.Rows[0][0].ToString();
                return ip;
            }
            catch
            {
                return null;
            }
        }
    }
}
