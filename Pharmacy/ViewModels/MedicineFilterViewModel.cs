using Microsoft.AspNetCore.Mvc.Rendering;
using Pharmacy.Models;

namespace Pharmacy.ViewModels;

public class MedicineFilterViewModel
{
    public string? Search { get; set; }
    public int? CategoryId { get; set; }
    public IEnumerable<Medicine> Medicines { get; set; } = [];
    public IEnumerable<SelectListItem> Categories { get; set; } = [];
}
