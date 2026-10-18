using Microsoft.AspNetCore.Mvc.Rendering;

namespace Pharmacy.ViewModels;

public class SaleCreateViewModel
{
    public List<SaleLineViewModel> Lines { get; set; } =
    [
        new(),
        new(),
        new()
    ];

    public IEnumerable<SelectListItem> Medicines { get; set; } = [];
}

public class SaleLineViewModel
{
    public int MedicineId { get; set; }
    public int Quantity { get; set; }
}
