using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Pharmacy.Models;

public class Sale
{
    public int Id { get; set; }

    [DataType(DataType.Date)]
    [Display(Name = "Дата продажи")]
    public DateOnly SaleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Column(TypeName = "numeric(12,2)")]
    [Display(Name = "Сумма")]
    public decimal TotalAmount { get; set; }

    public ICollection<SaleItem> Items { get; set; } = new List<SaleItem>();
}
