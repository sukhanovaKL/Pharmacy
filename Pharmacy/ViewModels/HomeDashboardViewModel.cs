using Pharmacy.Models;

namespace Pharmacy.ViewModels;

public class HomeDashboardViewModel
{
    public int MedicinesCount { get; set; }
    public int CategoriesCount { get; set; }
    public int SuppliersCount { get; set; }
    public int LowStockCount { get; set; }
    public List<Medicine> LowStockMedicines { get; set; } = [];
}
