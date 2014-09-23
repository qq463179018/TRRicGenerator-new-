using System.Data.Entity.ModelConfiguration.Conventions;

namespace Ric.Db.Model
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class HongKongModel : DbContext
    {
        public HongKongModel()
            : base("name=HongKongModel")
        {
            ETI_HK_StampDuties = Set<ETI_HK_StampDuty>();
            ETI_HK_TradingNews_ExlNames = Set<ETI_HK_TradingNews_ExlName>();
            ETI_HK_TradingNews_ExpireDates = Set<ETI_HK_TradingNews_ExpireDate>();
        }

        public virtual DbSet<ETI_HK_StampDuty> ETI_HK_StampDuties { get; set; }

        public virtual DbSet<ETI_HK_TradingNews_ExlName> ETI_HK_TradingNews_ExlNames { get; set; }

        public virtual DbSet<ETI_HK_TradingNews_ExpireDate> ETI_HK_TradingNews_ExpireDates { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            
            modelBuilder.Entity<ETI_HK_StampDuty>().ToTable("ETI_HK_StampDuty");
            modelBuilder.Entity<ETI_HK_TradingNews_ExlName>().ToTable("ETI_HK_TradingNews_ExlName");
            modelBuilder.Entity<ETI_HK_TradingNews_ExpireDate>().ToTable("ETI_HK_TradingNews_ExpireDate");

            modelBuilder.Entity<ETI_HK_StampDuty>()
                .Property(e => e.Ric)
                .IsFixedLength();

            modelBuilder.Entity<ETI_HK_StampDuty>()
                .Property(e => e.SubjectToStampDuty)
                .IsFixedLength();
        }
    }
}
