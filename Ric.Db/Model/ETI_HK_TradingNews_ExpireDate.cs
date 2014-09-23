namespace Ric.Db.Model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class ETI_HK_TradingNews_ExpireDate
    {
        public int Id { get; set; }

        [Required]
        [StringLength(50)]
        public string SourceDate { get; set; }

        [Required]
        [StringLength(50)]
        public string MappingDate { get; set; }

        public ETI_HK_TradingNews_ExpireDate()
        {
            SourceDate = "test";
            MappingDate = "test";
        }
    }
}
