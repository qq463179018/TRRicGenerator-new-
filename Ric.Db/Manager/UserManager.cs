using System;
using System.Collections.Generic;
using System.Linq;
using Ric.Db.Info;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class UserManager : ManagerBase
    {
        public static List<User> GetAllUsers(EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from user in ctx.Users
                        select user).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static List<User> GetUserNamesByManager(User manager, EtiRicGeneratorEntities ctx)
        {
            try
            {
                if (manager.Group == UserGroup.Admin)
                {
                    return (from user in ctx.Users
                            select user).ToList();
                }
                return (from user in ctx.Users
                        where user.ManagerId == manager.Id
                        select user).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static User GetUserByName(string name, EtiRicGeneratorEntities ctx)
        {
            return (from user in ctx.Users
                    where user.Familyname == name
                    select user).FirstOrDefault();
        }

        public static List<User> GetUserByGroup(UserGroup usergroup, EtiRicGeneratorEntities ctx)
        {
            return (from user in ctx.Users
                    where user.Group == usergroup
                    select user).ToList();
        }

        public static TaskInfo GetTaskByType(string typeFullName, EtiRicGeneratorEntities ctx)
        {
            try
            {
                var taskRes = (from task in ctx.Tasks
                                where task.GeneratorType.Contains(typeFullName)
                                select task).First();

                var task1 = new TaskInfo
                {
                    TaskId = taskRes.Id,
                    TaskName = taskRes.Name,
                    ConfigType = taskRes.ConfigType
                };
                return task1;
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

        public static List<Task> GetTaskByUsergroup(UserGroup userGroup, EtiRicGeneratorEntities ctx)
        {
            try
            {
                List<Task> result = null;
                if (userGroup == UserGroup.Dev)
                {
                    result = (from task in ctx.Tasks
                        select task).ToList();
                }
                if (userGroup == UserGroup.Admin)
                {
                    result = (from task in ctx.Tasks
                        where task.Status == TaskStatus.Active || task.Status == TaskStatus.Disabled
                        select task).ToList();
                }
                if (userGroup == UserGroup.User)
                {
                    result = (from task in ctx.Tasks
                        where task.Status == TaskStatus.Active
                        select task).ToList();
                }
                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
