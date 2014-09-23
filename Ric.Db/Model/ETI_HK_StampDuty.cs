namespace Ric.Db.Model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class ETI_HK_StampDuty
    {
        public int Id { get; set; }

        [Required]
        [StringLength(10)]
        public string Ric { get; set; }

        [Required]
        [StringLength(10)]
        public string SubjectToStampDuty { get; set; }

        [Column(TypeName = "date")]
        public DateTime LastChange { get; set; }
    }
}
