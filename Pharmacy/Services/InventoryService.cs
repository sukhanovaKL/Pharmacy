using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;

namespace Pharmacy.Services;

public class InventoryService(PharmacyDbContext db)
{
    public async Task CreateSupplyAsync(Supply supply, CancellationToken ct = default)
    {
        await using var tx = await db.Database.BeginTransactionAsync(ct);

        var medicine = await db.Medicines.FirstOrDefaultAsync(m => m.Id == supply.MedicineId, ct)
            ?? throw new InvalidOperationException("Препарат не найден.");

        db.Supplies.Add(supply);
        medicine.StockQuantity += supply.Quantity;
        await db.SaveChangesAsync(ct);
        await tx.CommitAsync(ct);
    }

    public async Task DeleteSupplyAsync(int supplyId, CancellationToken ct = default)
    {
        await using var tx = await db.Database.BeginTransactionAsync(ct);

        var supply = await db.Supplies.FirstOrDefaultAsync(s => s.Id == supplyId, ct)
            ?? throw new InvalidOperationException("Поступление не найдено.");

        var medicine = await db.Medicines.FirstOrDefaultAsync(m => m.Id == supply.MedicineId, ct)
            ?? throw new InvalidOperationException("Препарат не найден.");

        if (medicine.StockQuantity < supply.Quantity)
        {
            throw new InvalidOperationException(
                "Нельзя удалить поступление: текущий остаток меньше количества в поступлении.");
        }

        medicine.StockQuantity -= supply.Quantity;
        db.Supplies.Remove(supply);
        await db.SaveChangesAsync(ct);
        await tx.CommitAsync(ct);
    }

    public async Task CreateSaleAsync(IReadOnlyList<(int MedicineId, int Quantity)> lines, CancellationToken ct = default)
    {
        if (lines.Count == 0)
            throw new InvalidOperationException("Добавьте хотя бы один препарат.");

        await using var tx = await db.Database.BeginTransactionAsync(ct);

        var sale = new Sale
        {
            SaleDate = DateOnly.FromDateTime(DateTime.Today),
            TotalAmount = 0
        };

        foreach (var (medicineId, quantity) in lines)
        {
            if (quantity <= 0)
                throw new InvalidOperationException("Количество должно быть больше 0.");

            var medicine = await db.Medicines.FirstOrDefaultAsync(m => m.Id == medicineId, ct)
                ?? throw new InvalidOperationException($"Препарат #{medicineId} не найден.");

            if (medicine.StockQuantity < quantity)
            {
                throw new InvalidOperationException(
                    $"Недостаточно «{medicine.Name}»: в наличии {medicine.StockQuantity}, запрошено {quantity}.");
            }

            medicine.StockQuantity -= quantity;
            sale.Items.Add(new SaleItem
            {
                MedicineId = medicine.Id,
                Quantity = quantity,
                UnitPrice = medicine.Price
            });
            sale.TotalAmount += medicine.Price * quantity;
        }

        db.Sales.Add(sale);
        await db.SaveChangesAsync(ct);
        await tx.CommitAsync(ct);
    }

    public async Task DeleteSaleAsync(int saleId, CancellationToken ct = default)
    {
        await using var tx = await db.Database.BeginTransactionAsync(ct);

        var sale = await db.Sales
            .Include(s => s.Items)
            .FirstOrDefaultAsync(s => s.Id == saleId, ct)
            ?? throw new InvalidOperationException("Продажа не найдена.");

        foreach (var item in sale.Items)
        {
            var medicine = await db.Medicines.FirstOrDefaultAsync(m => m.Id == item.MedicineId, ct)
                ?? throw new InvalidOperationException("Препарат не найден.");
            medicine.StockQuantity += item.Quantity;
        }

        db.Sales.Remove(sale);
        await db.SaveChangesAsync(ct);
        await tx.CommitAsync(ct);
    }
}
