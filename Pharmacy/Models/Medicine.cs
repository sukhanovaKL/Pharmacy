using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Pharmacy.Models;

public class Medicine
{
    public int Id { get; set; }

    [Required(ErrorMessage = "Укажите название препарата")]
    [StringLength(150)]
    [Display(Name = "Название")]
    public string Name { get; set; } = string.Empty;

    [Required(ErrorMessage = "Укажите производителя")]
    [StringLength(150)]
    [Display(Name = "Производитель")]
    public string Manufacturer { get; set; } = string.Empty;

    [Range(0.01, 1000000, ErrorMessage = "Цена должна быть больше 0")]
    [Column(TypeName = "numeric(12,2)")]
    [Display(Name = "Цена")]
    public decimal Price { get; set; }

    [Range(0, int.MaxValue, ErrorMessage = "Остаток не может быть отрицательным")]
    [Display(Name = "Остаток")]
    public int StockQuantity { get; set; }

    [Display(Name = "Нужен рецепт")]
    public bool RequiresPrescription { get; set; }

    [Display(Name = "Категория")]
    public int CategoryId { get; set; }

    public Category? Category { get; set; }

    public ICollection<Supply> Supplies { get; set; } = new List<Supply>();
    public ICollection<SaleItem> SaleItems { get; set; } = new List<SaleItem>();
}
