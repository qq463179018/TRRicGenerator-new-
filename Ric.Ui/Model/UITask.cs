using System;
using Ric.Db.Info;
using Ric.Db.Model;

namespace Ric.Ui.Model
{
    public class UITask
    {
        public UITask()
        {
            TaskId = TaskName = Market = string.Empty;
            Selected = false;
            ConfigObjectType = null;
            //Status = TaskStatus.;
            CostTime = "00:00:00:00";
            //TaskResultList = null;
        }

        public string TaskId { get; set; }
        public bool Selected { get; set; }
        public string TaskName { get; set; }
        public string Market { get; set; }
        public Type ConfigObjectType { get; set; }
        public Type GeneratorObjectType { get; set; }
        public TaskStatus Status { get; set; }
        public string CostTime { get; set; }
        public string GroupName { get; set; }
        public string Description { get; set; }
        public string DisplayId { get; set; }
        //public Image Icon
        //{
        //    get
        //    {
        //        switch (Status)
        //        {
        //            case TaskStatus.Completed:
        //                return global::Ric.Generator.UI.Properties.Resources.Completed;
        //            case TaskStatus.Failed:
        //                return global::Ric.Generator.UI.Properties.Resources.Failed;
        //            case TaskStatus.Ready:
        //                return global::Ric.Generator.UI.Properties.Resources.Ready;
        //            case TaskStatus.Running:
        //            default:
        //                return global::Ric.Generator.UI.Properties.Resources.Running;
        //        }
        //    }
        //}

        //public List<TaskResultEntry> TaskResultList { get; set; }
    }
}