namespace WebApp.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class AppDBContext : DbContext
    {
        public AppDBContext()
            : base("name=DefaultConnection")
        {
        }

        public virtual DbSet<EDI_754_945_ShipToChange> EDI_754_945_ShipToChange { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
