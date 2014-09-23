using System.Linq;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class TaskConfigManager : ManagerBase
    {
        public static string GetConfigValue(int userId, int taskId, string configName, EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from config in ctx.Configs
                        where config.UserId == userId
                              && config.TaskId == taskId
                              && config.Key == configName
                        select config.Value).Single();
            }
            catch
            {
                return "";
            }
        }

        public static void UpdateConfig(int userId, int taskId, string configName, string configValue, EtiRicGeneratorEntities ctx)
        {
            try
            {
                var configVal = (from config in ctx.Configs
                                where config.UserId == userId
                                      && config.TaskId == taskId
                                      && config.Key == configName
                                select config).Single();

                if (configVal.Value != configValue)
                {
                    configVal.Value = configValue;
                    ctx.SaveChanges();
                }
            }
            catch
            {
                ctx.Configs.Add(new Model.Config
                {
                    TaskId = taskId,
                    UserId = userId,
                    Key = configName,
                    Value = configValue
                });
                ctx.SaveChanges();
            }
        }
    }
}