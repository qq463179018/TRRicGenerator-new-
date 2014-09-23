using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Ric.Db.Model;

namespace Ric.Db.Manager
{
    public class ScheduleManager
    {
        public static List<Schedule> GetScheduledTask(User user, EtiRicGeneratorEntities ctx)
        {
            return (from schedule in ctx.Schedules
                    where schedule.UserId == user.Id
                    select schedule).ToList();
        }

    }
}
