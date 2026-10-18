using System.ComponentModel.DataAnnotations;

namespace Pharmacy.Models;

public class Supply
{
    public int Id { get; set; }

    [Display(Name = "Препарат")]
    public int MedicineId { get; set; }

    public Medicine? Medicine { get; set; }

    [Display(Name = "Поставщик")]
    public int SupplierId { get; set; }

    public Supplier? Supplier { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Количество должно быть не меньше 1")]
    [Display(Name = "Количество")]
    public int Quantity { get; set; }

    [DataType(DataType.Date)]
    [Display(Name = "Дата поступления")]
    public DateOnly SupplyDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
}
