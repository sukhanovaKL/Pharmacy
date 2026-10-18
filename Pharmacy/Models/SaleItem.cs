using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Pharmacy.Models;

public class SaleItem
{
    public int Id { get; set; }

    public int SaleId { get; set; }
    public Sale? Sale { get; set; }

    [Display(Name = "Препарат")]
    public int MedicineId { get; set; }

    public Medicine? Medicine { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Количество должно быть не меньше 1")]
    [Display(Name = "Количество")]
    public int Quantity { get; set; }

    [Column(TypeName = "numeric(12,2)")]
    [Display(Name = "Цена за ед.")]
    public decimal UnitPrice { get; set; }
}
