using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;

namespace Pharmacy.Controllers;

public class SuppliersController(PharmacyDbContext db) : Controller
{
    public async Task<IActionResult> Index(string? search)
    {
        ViewBag.Search = search;
        var query = db.Suppliers.AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
        {
            query = query.Where(s =>
                s.Name.Contains(search) ||
                (s.Phone != null && s.Phone.Contains(search)) ||
                (s.Address != null && s.Address.Contains(search)));
        }

        return View(await query.OrderBy(s => s.Name).ToListAsync());
    }

    public IActionResult Create() => View(new Supplier());

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(Supplier supplier)
    {
        if (!ModelState.IsValid)
            return View(supplier);

        db.Suppliers.Add(supplier);
        try
        {
            await db.SaveChangesAsync();
        }
        catch
        {
            ModelState.AddModelError(string.Empty, "Не удалось сохранить. Возможно, такое название уже есть.");
            return View(supplier);
        }

        TempData["Success"] = "Поставщик создан.";
        return RedirectToAction(nameof(Index));
    }

    public async Task<IActionResult> Edit(int id)
    {
        var supplier = await db.Suppliers.FindAsync(id);
        return supplier is null ? NotFound() : View(supplier);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(int id, Supplier supplier)
    {
        if (id != supplier.Id)
            return NotFound();

        if (!ModelState.IsValid)
            return View(supplier);

        db.Attach(supplier).State = EntityState.Modified;
        try
        {
            await db.SaveChangesAsync();
        }
        catch (DbUpdateException)
        {
            ModelState.AddModelError(string.Empty, "Не удалось сохранить. Возможно, такое название уже есть.");
            return View(supplier);
        }

        TempData["Success"] = "Поставщик обновлён.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        var supplier = await db.Suppliers.Include(s => s.Supplies).FirstOrDefaultAsync(s => s.Id == id);
        if (supplier is null)
        {
            TempData["Error"] = "Поставщик не найден.";
            return RedirectToAction(nameof(Index));
        }

        if (supplier.Supplies.Count > 0)
        {
            TempData["Error"] = "Нельзя удалить поставщика: есть связанные поступления.";
            return RedirectToAction(nameof(Index));
        }

        db.Suppliers.Remove(supplier);
        await db.SaveChangesAsync();
        TempData["Success"] = "Поставщик удалён.";
        return RedirectToAction(nameof(Index));
    }
}
