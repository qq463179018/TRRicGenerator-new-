
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
    
public partial class Market
{

    public Market()
    {

        this.Tasks = new HashSet<Task>();

        this.Users = new HashSet<User>();

    }


    public int Id { get; set; }

    public string Name { get; set; }

    public string Abbreviation { get; set; }

    public int ManagerId { get; set; }



    public virtual ICollection<Task> Tasks { get; set; }

    public virtual User Manager { get; set; }

    public virtual ICollection<User> Users { get; set; }

}

}
