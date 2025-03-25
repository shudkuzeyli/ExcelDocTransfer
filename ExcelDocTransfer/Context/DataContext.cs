using ExcelDocTransfer.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelDocTransfer.Context
{
	public class DataContext : DbContext
	{
		public DataContext(DbContextOptions<DataContext> options) : base(options)
		{
		}
		public DataContext()
		{

		}

		public DbSet<CustomerResponse> CustomerResponses { get; set; }


		protected override void OnModelCreating(ModelBuilder modelBuilder)
		{
			modelBuilder.Entity<CustomerResponse>(entity =>
			{
				entity.HasKey(e => e.Id)
				.HasName("PK_CustomerResponses");

				entity.Property(e => e.CustomerName)
				.IsRequired()
					.HasMaxLength(50)
					.IsUnicode(true);

				entity.Property(e => e.CustomerAddress)
				.IsRequired()
					.HasMaxLength(50)
					.IsUnicode(true);

				entity.Property(e => e.CustomerCity)
				.IsRequired()
					.HasMaxLength(50)
					.IsUnicode(true);

				entity.Property(e => e.Fees)
				.HasColumnType("decimal(18, 2)");

				entity.Property(e => e.VisitDate)
				.HasColumnType("datetime");
			});

			base.OnModelCreating(modelBuilder);
		}
		//partial void OnModelCreatingPartial(ModelBuilder modelBuilder);

		protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
		{
			//veritabanı bağlantı ayarları appsettings .json dosyasında tanımlanmışsa.
			if (!optionsBuilder.IsConfigured)
			{
				optionsBuilder.UseSqlServer("Server=.;Database=ExcelDocTransfer;Trusted_Connection=True;");
			}
		}
	}
}
