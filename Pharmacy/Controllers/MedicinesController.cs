using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;
using Pharmacy.ViewModels;

namespace Pharmacy.Controllers;

public class MedicinesController(PharmacyDbContext db) : Controller
{
    public async Task<IActionResult> Index(string? search, int? categoryId)
    {
        var query = db.Medicines.Include(m => m.Category).AsQueryable();

        if (!string.IsNullOrWhiteSpace(search))
        {
            query = query.Where(m =>
                m.Name.Contains(search) ||
                m.Manufacturer.Contains(search));
        }

        if (categoryId is not null)
            query = query.Where(m => m.CategoryId == categoryId);

        var model = new MedicineFilterViewModel
        {
            Search = search,
            CategoryId = categoryId,
            Medicines = await query.OrderBy(m => m.Name).ToListAsync(),
            Categories = await db.Categories
                .OrderBy(c => c.Name)
                .Select(c => new SelectListItem { Value = c.Id.ToString(), Text = c.Name })
                .ToListAsync()
        };

        return View(model);
    }

    public async Task<IActionResult> Create()
    {
        await LoadCategoriesAsync();
        return View(new Medicine());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(Medicine medicine)
    {
        await LoadCategoriesAsync();
        if (!ModelState.IsValid)
            return View(medicine);

        db.Medicines.Add(medicine);
        await db.SaveChangesAsync();
        TempData["Success"] = "Препарат создан.";
        return RedirectToAction(nameof(Index));
    }

    public async Task<IActionResult> Edit(int id)
    {
        var medicine = await db.Medicines.FindAsync(id);
        if (medicine is null)
            return NotFound();

        await LoadCategoriesAsync();
        return View(medicine);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(int id, Medicine medicine)
    {
        if (id != medicine.Id)
            return NotFound();

        await LoadCategoriesAsync();
        if (!ModelState.IsValid)
            return View(medicine);

        db.Attach(medicine).State = EntityState.Modified;
        await db.SaveChangesAsync();
        TempData["Success"] = "Препарат обновлён.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        var medicine = await db.Medicines
            .Include(m => m.Supplies)
            .Include(m => m.SaleItems)
            .FirstOrDefaultAsync(m => m.Id == id);

        if (medicine is null)
        {
            TempData["Error"] = "Препарат не найден.";
            return RedirectToAction(nameof(Index));
        }

        if (medicine.Supplies.Count > 0 || medicine.SaleItems.Count > 0)
        {
            TempData["Error"] = "Нельзя удалить препарат: есть связанные поступления или продажи.";
            return RedirectToAction(nameof(Index));
        }

        db.Medicines.Remove(medicine);
        await db.SaveChangesAsync();
        TempData["Success"] = "Препарат удалён.";
        return RedirectToAction(nameof(Index));
    }

    private async Task LoadCategoriesAsync()
    {
        ViewBag.Categories = new SelectList(
            await db.Categories.OrderBy(c => c.Name).ToListAsync(),
            "Id",
            "Name");
    }
}
