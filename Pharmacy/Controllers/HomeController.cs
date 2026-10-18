using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Pharmacy.Data;
using Pharmacy.Models;
using Pharmacy.ViewModels;

namespace Pharmacy.Controllers;

public class HomeController(PharmacyDbContext db) : Controller
{
    public async Task<IActionResult> Index()
    {
        var model = new HomeDashboardViewModel
        {
            MedicinesCount = await db.Medicines.CountAsync(),
            CategoriesCount = await db.Categories.CountAsync(),
            SuppliersCount = await db.Suppliers.CountAsync(),
            LowStockCount = await db.Medicines.CountAsync(m => m.StockQuantity < 10),
            LowStockMedicines = await db.Medicines
                .Include(m => m.Category)
                .Where(m => m.StockQuantity < 10)
                .OrderBy(m => m.StockQuantity)
                .Take(10)
                .ToListAsync()
        };

        return View(model);
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel
        {
            RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
        });
    }
}

public class ErrorViewModel
{
    public string? RequestId { get; set; }
    public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
}
