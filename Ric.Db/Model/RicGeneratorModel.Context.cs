﻿

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
using System.Data.Entity;
using System.Data.Entity.Infrastructure;


public partial class EtiRicGeneratorEntities : DbContext
{
    public EtiRicGeneratorEntities()
        : base("name=EtiRicGeneratorEntities")
    {

    }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        throw new UnintentionalCodeFirstException();
    }


    public virtual DbSet<Config> Configs { get; set; }

    public virtual DbSet<Market> Markets { get; set; }

    public virtual DbSet<Run> Runs { get; set; }

    public virtual DbSet<Schedule> Schedules { get; set; }

    public virtual DbSet<Task> Tasks { get; set; }

    public virtual DbSet<User> Users { get; set; }

}

}
