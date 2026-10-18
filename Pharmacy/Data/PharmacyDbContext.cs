using Microsoft.EntityFrameworkCore;
using Pharmacy.Models;

namespace Pharmacy.Data;

public class PharmacyDbContext(DbContextOptions<PharmacyDbContext> options) : DbContext(options)
{
    public DbSet<Category> Categories => Set<Category>();
    public DbSet<Supplier> Suppliers => Set<Supplier>();
    public DbSet<Medicine> Medicines => Set<Medicine>();
    public DbSet<Supply> Supplies => Set<Supply>();
    public DbSet<Sale> Sales => Set<Sale>();
    public DbSet<SaleItem> SaleItems => Set<SaleItem>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Category>()
            .HasIndex(c => c.Name)
            .IsUnique();

        modelBuilder.Entity<Supplier>()
            .HasIndex(s => s.Name)
            .IsUnique();

        modelBuilder.Entity<Medicine>()
            .HasOne(m => m.Category)
            .WithMany(c => c.Medicines)
            .HasForeignKey(m => m.CategoryId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<Supply>()
            .HasOne(s => s.Medicine)
            .WithMany(m => m.Supplies)
            .HasForeignKey(s => s.MedicineId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<Supply>()
            .HasOne(s => s.Supplier)
            .WithMany(s => s.Supplies)
            .HasForeignKey(s => s.SupplierId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<SaleItem>()
            .HasOne(i => i.Sale)
            .WithMany(s => s.Items)
            .HasForeignKey(i => i.SaleId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<SaleItem>()
            .HasOne(i => i.Medicine)
            .WithMany(m => m.SaleItems)
            .HasForeignKey(i => i.MedicineId)
            .OnDelete(DeleteBehavior.Restrict);

        SeedData(modelBuilder);
    }

    private static void SeedData(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Category>().HasData(
            new Category { Id = 1, Name = "Жаропонижающие" },
            new Category { Id = 2, Name = "Антибиотики" },
            new Category { Id = 3, Name = "Витамины" }
        );

        modelBuilder.Entity<Supplier>().HasData(
            new Supplier { Id = 1, Name = "ФармОпт", Phone = "+7 (495) 111-22-33", Address = "г. Москва, ул. Ленина, 10" },
            new Supplier { Id = 2, Name = "МедСнаб", Phone = "+7 (812) 444-55-66", Address = "г. Санкт-Петербург, Невский пр., 25" }
        );

        modelBuilder.Entity<Medicine>().HasData(
            new Medicine
            {
                Id = 1,
                Name = "Парацетамол",
                Manufacturer = "Фармстандарт",
                Price = 85.50m,
                StockQuantity = 120,
                RequiresPrescription = false,
                CategoryId = 1
            },
            new Medicine
            {
                Id = 2,
                Name = "Амоксициллин",
                Manufacturer = "Синтез",
                Price = 210.00m,
                StockQuantity = 45,
                RequiresPrescription = true,
                CategoryId = 2
            },
            new Medicine
            {
                Id = 3,
                Name = "Аскорбиновая кислота",
                Manufacturer = "Марбиофарм",
                Price = 45.00m,
                StockQuantity = 8,
                RequiresPrescription = false,
                CategoryId = 3
            }
        );
    }
}
