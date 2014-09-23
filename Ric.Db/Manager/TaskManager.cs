using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Ric.Db.Info;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class TaskManager : ManagerBase
    {
        public static List<Task> GetAllTasks(EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from task in ctx.Tasks
                    select task).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static List<string> GetAllTasksName(EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from task in ctx.Tasks
                    select task.Name).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static TaskInfo GetTaskByType(string typeFullName, EtiRicGeneratorEntities ctx)
        {
            try
            {
                var taskRes = (from task in ctx.Tasks
                    where task.GeneratorType.Contains(typeFullName)
                    select task).First();

                return new TaskInfo
                {
                    TaskId = taskRes.Id,
                    TaskName = taskRes.Name,
                    ConfigType = taskRes.ConfigType
                };
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static Task GetTaskByName(string taskName, EtiRicGeneratorEntities ctx)
        {
            return (from task in ctx.Tasks
                where task.Name == taskName
                select task).FirstOrDefault();
        }


        public static List<Task> GetTaskByGroupMarket(User user, EtiRicGeneratorEntities ctx)
        {
            try
            {
                if (user.Group == UserGroup.Dev)
                {
                    return (from task in ctx.Tasks
                            where task.MarketId == user.MainMarketId
                            select task).ToList();
                }
                if (user.Group == UserGroup.Admin)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active || task.Status == TaskStatus.Disabled
                            && task.MarketId == user.MainMarketId
                            select task).ToList();
                }
                if (user.Group == UserGroup.Manager)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active
                            && task.MarketId == user.MainMarketId
                            select task).ToList();
                }
                if (user.Group == UserGroup.User)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active
                            && task.MarketId == user.MainMarketId
                            select task).ToList();
                }
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static List<Task> GetTaskByGroup(User user, EtiRicGeneratorEntities ctx)
        {
            try
            {
#if DEBUG
                return (from task in ctx.Tasks
                        select task).ToList();
#else

                if (user.Group == UserGroup.Dev)
                {
                    return (from task in ctx.Tasks
                            select task).ToList();
                }
                if (user.Group == UserGroup.Admin)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active || task.Status == TaskStatus.Disabled
                            select task).ToList();
                }
                if (user.Group == UserGroup.Manager)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active
                            select task).ToList();
                }
                if (user.Group == UserGroup.User)
                {
                    return (from task in ctx.Tasks
                            where task.Status == TaskStatus.Active
                            select task).ToList();
                }
                return null;
#endif
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static IQueryable<IGrouping<string, Task>> GetTaskTreeByGroupMarket(User user, EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from task in ctx.Tasks
                        join market in ctx.Markets on task.MarketId equals market.Id
                        orderby task.Name 
                        group task by market.Name
                        into marketlist
                        select marketlist);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static Task GetTaskById(int? taskId, EtiRicGeneratorEntities ctx)
        {
            return (from task in ctx.Tasks
                    where task.Id == taskId
                    select task).FirstOrDefault();
        }

        public static List<Task> GetTasksByManager(User user, EtiRicGeneratorEntities ctx)
        {
            var tasksToReturn = new List<Task>();

            foreach (var market in user.ManagingMarkets)
            {
                tasksToReturn.AddRange(market.Tasks);
            }
            return tasksToReturn;
        }

        public static int GetTimeSaved(Task newTask, int daysToAdd, EtiRicGeneratorEntities ctx)
        {
            var toComparedate = DateTime.Now.AddDays(daysToAdd);
            return ((from run in ctx.Runs
                where run.TaskId == newTask.Id
                      && run.Result == TaskResult.Success
                      && run.Date.Value.CompareTo(toComparedate) > 0
                select run).Count()
                         * newTask.ManualTime);

        }

        public static double AverageRunningTime(Task task, EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from run in ctx.Runs
                        where run.TaskId == task.Id
                        select run.Duration).Average();
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public static IEnumerable<Task> GetTaskByOwner(User currentUser, EtiRicGeneratorEntities ctx)
        {
            return (from task in ctx.Tasks
                    where task.OwnerId == currentUser.Id
                    select task);
        }

        public static IEnumerable<Task> GetTaskByMarket(int marketId, EtiRicGeneratorEntities ctx)
        {
            return (from task in ctx.Tasks
                    where task.MarketId == marketId
                    select task);
        }
    }
}
