using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;

namespace Pharmacy.Controllers;

public class CategoriesController(PharmacyDbContext db) : Controller
{
    public async Task<IActionResult> Index(string? search)
    {
        ViewBag.Search = search;
        var query = db.Categories.AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
            query = query.Where(c => c.Name.Contains(search));

        return View(await query.OrderBy(c => c.Name).ToListAsync());
    }

    public IActionResult Create() => View(new Category());

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(Category category)
    {
        if (!ModelState.IsValid)
            return View(category);

        db.Categories.Add(category);
        try
        {
            await db.SaveChangesAsync();
        }
        catch
        {
            ModelState.AddModelError(string.Empty, "Не удалось сохранить. Возможно, такое название уже есть.");
            return View(category);
        }

        TempData["Success"] = "Категория создана.";
        return RedirectToAction(nameof(Index));
    }

    public async Task<IActionResult> Edit(int id)
    {
        var category = await db.Categories.FindAsync(id);
        return category is null ? NotFound() : View(category);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(int id, Category category)
    {
        if (id != category.Id)
            return NotFound();

        if (!ModelState.IsValid)
            return View(category);

        db.Attach(category).State = EntityState.Modified;
        try
        {
            await db.SaveChangesAsync();
        }
        catch (DbUpdateException)
        {
            ModelState.AddModelError(string.Empty, "Не удалось сохранить. Возможно, такое название уже есть.");
            return View(category);
        }

        TempData["Success"] = "Категория обновлена.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        var category = await db.Categories.Include(c => c.Medicines).FirstOrDefaultAsync(c => c.Id == id);
        if (category is null)
        {
            TempData["Error"] = "Категория не найдена.";
            return RedirectToAction(nameof(Index));
        }

        if (category.Medicines.Count > 0)
        {
            TempData["Error"] = "Нельзя удалить категорию: есть связанные препараты.";
            return RedirectToAction(nameof(Index));
        }

        db.Categories.Remove(category);
        await db.SaveChangesAsync();
        TempData["Success"] = "Категория удалена.";
        return RedirectToAction(nameof(Index));
    }
}
