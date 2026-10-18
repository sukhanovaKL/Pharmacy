using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;
using Pharmacy.Services;

namespace Pharmacy.Controllers;

public class SuppliesController(PharmacyDbContext db, InventoryService inventory) : Controller
{
    public async Task<IActionResult> Index()
    {
        var supplies = await db.Supplies
            .Include(s => s.Medicine)
            .Include(s => s.Supplier)
            .OrderByDescending(s => s.SupplyDate)
            .ThenByDescending(s => s.Id)
            .ToListAsync();

        return View(supplies);
    }

    public async Task<IActionResult> Create(int? medicineId)
    {
        await LoadLookupsAsync();
        return View(new Supply
        {
            MedicineId = medicineId ?? 0,
            SupplyDate = DateOnly.FromDateTime(DateTime.Today),
            Quantity = 1
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(Supply supply)
    {
        await LoadLookupsAsync();
        if (!ModelState.IsValid)
            return View(supply);

        try
        {
            await inventory.CreateSupplyAsync(supply);
        }
        catch (InvalidOperationException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(supply);
        }

        TempData["Success"] = "Поступление оформлено, остаток увеличен.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        try
        {
            await inventory.DeleteSupplyAsync(id);
            TempData["Success"] = "Поступление удалено, остаток обновлён.";
        }
        catch (InvalidOperationException ex)
        {
            TempData["Error"] = ex.Message;
        }

        return RedirectToAction(nameof(Index));
    }

    private async Task LoadLookupsAsync()
    {
        ViewBag.Medicines = new SelectList(await db.Medicines.OrderBy(m => m.Name).ToListAsync(), "Id", "Name");
        ViewBag.Suppliers = new SelectList(await db.Suppliers.OrderBy(s => s.Name).ToListAsync(), "Id", "Name");
    }
}
