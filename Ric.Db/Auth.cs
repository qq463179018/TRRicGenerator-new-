using System;
using System.Linq;
using Ric.Db.Manager;
using Ric.Db.Model;

namespace Ric.Db
{
    public class Auth : ManagerBase
    {
        public static User GetEtiUser(EtiRicGeneratorEntities ctx)
        {
            return GetEtiUser(GetCurrentWinUser(), ctx);
        }

        public static User GetAnonymousUser()
        {
            return new User
            {
                Id = 2,
                WinUser = "Anonymous",
                GedaUser = "ETIASIA",
                GedaPassword = "ETIASIA",
                Group = UserGroup.User,
                Status = UserStatus.Active,
            };
        }

        public static User GetEtiUser(string winUser, EtiRicGeneratorEntities ctx)
        {
            return (from user in ctx.Users
                where user.WinUser == winUser
                select user).ToList().Single();
        }

        public static User GetEtiUserById(int userId, EtiRicGeneratorEntities ctx)
        {
            try
            {
                return (from user in ctx.Users
                    where user.Id == userId
                    select user).Single();
            }
            catch
            {
                return null;
            }
        }

        public static void AddEtiUser(EtiRicGeneratorEntities ctx)
        {
            ctx.Users.Add(new User
            {
                WinUser = GetCurrentWinUser(),
                GedaUser = "ETIASIA",
                GedaPassword = "ETIASIA",
                Group = UserGroup.User,
                Status = UserStatus.Active,
            });
            ctx.SaveChanges();
        }

        public static string GetCurrentWinUser()
        {
            return Environment.UserName;
        }

    }
}
