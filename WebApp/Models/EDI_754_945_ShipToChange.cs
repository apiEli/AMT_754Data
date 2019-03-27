namespace WebApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class EDI_754_945_ShipToChange
    {
        [Key]
        public int ID { get; set; }

        [StringLength(255)]
        public string CompanyCode { get; set; }

        [StringLength(255)]
        public string DivisionCode { get; set; }

        [StringLength(255)]
        public string CustomerNumber { get; set; }

        [Column("RRC#")]
        [StringLength(255)]
        public string RRC_ { get; set; }

        [Column("Load ID")]
        [StringLength(255)]
        public string Load_ID { get; set; }

        [StringLength(255)]
        public string SCAC { get; set; }

        [Column("Service Level")]
        [StringLength(255)]
        public string Service_Level { get; set; }

        [Column("Catalog PO/Retail DI")]
        [StringLength(255)]
        public string Catalog_PO_Retail_DI { get; set; }

        [Column("Ship Date")]
        [StringLength(255)]
        public string Ship_Date { get; set; }

        [StringLength(255)]
        public string Destination { get; set; }
    }
}
