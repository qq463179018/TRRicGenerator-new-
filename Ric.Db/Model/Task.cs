
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------


namespace Ric.Db.Model
{

using System;
    using System.Collections.Generic;
    
public partial class Task
{

    public Task()
    {

        this.Schedules = new HashSet<Schedule>();

        this.Runs = new HashSet<Run>();

        this.Configs = new HashSet<Config>();

    }


    public int Id { get; set; }

    public string Name { get; set; }

    public string Group { get; set; }

    public TaskStatus Status { get; set; }

    public string Description { get; set; }

    public string ConfigType { get; set; }

    public string GeneratorType { get; set; }

    public int MarketId { get; set; }

    public int OwnerId { get; set; }

    public int ManualTime { get; set; }



    public virtual Market Market { get; set; }

    public virtual User Owner { get; set; }

    public virtual ICollection<Schedule> Schedules { get; set; }

    public virtual ICollection<Run> Runs { get; set; }

    public virtual ICollection<Config> Configs { get; set; }

}

}
