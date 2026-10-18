using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Services;
using Pharmacy.ViewModels;

namespace Pharmacy.Controllers;

public class SalesController(PharmacyDbContext db, InventoryService inventory) : Controller
{
    public async Task<IActionResult> Index()
    {
        var sales = await db.Sales
            .Include(s => s.Items)
            .ThenInclude(i => i.Medicine)
            .OrderByDescending(s => s.SaleDate)
            .ThenByDescending(s => s.Id)
            .ToListAsync();

        return View(sales);
    }

    public async Task<IActionResult> Details(int id)
    {
        var sale = await db.Sales
            .Include(s => s.Items)
            .ThenInclude(i => i.Medicine)
            .FirstOrDefaultAsync(s => s.Id == id);

        return sale is null ? NotFound() : View(sale);
    }

    public async Task<IActionResult> Create()
    {
        return View(await BuildCreateModelAsync());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(SaleCreateViewModel model)
    {
        var selected = model.Lines
            .Where(l => l.MedicineId > 0 && l.Quantity > 0)
            .Select(l => (l.MedicineId, l.Quantity))
            .ToList();

        if (selected.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Добавьте хотя бы одну позицию с препаратом и количеством.");
            return View(await BuildCreateModelAsync(model));
        }

        try
        {
            await inventory.CreateSaleAsync(selected);
        }
        catch (InvalidOperationException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(await BuildCreateModelAsync(model));
        }

        TempData["Success"] = "Продажа оформлена, остаток списан.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        try
        {
            await inventory.DeleteSaleAsync(id);
            TempData["Success"] = "Продажа удалена, остаток возвращён на склад.";
        }
        catch (InvalidOperationException ex)
        {
            TempData["Error"] = ex.Message;
        }

        return RedirectToAction(nameof(Index));
    }

    private async Task<SaleCreateViewModel> BuildCreateModelAsync(SaleCreateViewModel? model = null)
    {
        model ??= new SaleCreateViewModel();
        model.Medicines = await db.Medicines
            .OrderBy(m => m.Name)
            .Select(m => new SelectListItem
            {
                Value = m.Id.ToString(),
                Text = $"{m.Name} (остаток: {m.StockQuantity}, {m.Price:N2} ₽)"
            })
            .ToListAsync();
        return model;
    }
}
