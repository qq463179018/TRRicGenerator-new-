using System;
using System.Collections.Generic;
using Ric.Db.Model;
using User = Ric.Db.Model.User;

namespace Ric.Core
{
    public class ScheduleSetting
    {
        public ScheduleType Type { get; set; }
        public DateTimeOffset StartTime { get; set; }
        public string CronExpression { get; set; }
        public int RetryCount { get; set; }
        public int RetryIntervalInMinute { get; set; }
        public int RepeatCount { get; set; }
        public int RepeatIntervalInMinute { get; set; }
    }

    public enum ScheduleType
    {
        Workday,
        Daily,
        RunOnce,
        Repeat
    }

    public class ConfigStore
    {
        private Dictionary<int, object> configDict = new Dictionary<int, object>();

        public object GetConfig(int taskId)
        {
            return configDict.ContainsKey(taskId) ? configDict[taskId] : null;
        }

        public void StoreConfig(int taskId, object config)
        {
            if (!configDict.ContainsKey(taskId))
            {
                configDict.Add(taskId, config);
                return;
            }

            configDict[taskId] = config;
        }
    }

    public class RunTimeContext
    {
        private static RunTimeContext instance;
        private static object lockObj = new object();
        private ConfigStore configStore;

        private RunTimeContext()
        {
        }

        //public EtiRicGeneratorEntities DatabaseContext { get; set; }
        public EtiRicGeneratorEntities DatabaseContext { get; set; }

        public static RunTimeContext Context
        {
            get
            {
                if (instance == null)
                {
                    lock (lockObj)
                    {
                        if (instance == null)
                        {
                            instance = new RunTimeContext();
                        }
                    }
                }

                return instance;
            }
        }

        public ConfigStore ConfigStore
        {
            get { return configStore ?? (configStore = new ConfigStore()); }
        }

        public User CurrentUser { get; set; }
    }
}
