using System;
using System.Collections.Generic;
using System.Linq;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class StatsManager
    {
        public static void AddRunInfo(Run newRun, EtiRicGeneratorEntities ctx)
        {
            ctx.Runs.Add(newRun);
            ctx.SaveChanges();
        }

        public static List<Run> GetRunsByTask(int taskId, EtiRicGeneratorEntities ctx)
        {
            return (from run in ctx.Runs
                    where run.TaskId == taskId
                    orderby run.Date descending 
                    select run).ToList();
        }

        public static List<Run> GetRunsByUser(int userId, int nbToTake, EtiRicGeneratorEntities ctx)
        {
            return (from run in ctx.Runs
                    where run.UserId == userId
                    orderby run.Date descending 
                    select run).Take(nbToTake).ToList();
        }

        public static List<Run> GetSuccessfullRunsByTask(int taskId, int nbToTake, EtiRicGeneratorEntities ctx)
        {
            return (from run in ctx.Runs
                    where run.TaskId == taskId
                        && run.Result == TaskResult.Success
                    orderby run.Date descending
                    select run).Take(nbToTake).ToList();
        }

        public static int GetFailedRuns(Task newTask, EtiRicGeneratorEntities ctx)
        {
            var toCompareDate = DateTime.Now.AddDays(-7);
            return (from run in ctx.Runs
                    where run.TaskId == newTask.Id
                        && run.Date.Value.CompareTo(toCompareDate) > 0
                        && run.Result == TaskResult.Fail
                    select run).Count();
        }
    }
}
