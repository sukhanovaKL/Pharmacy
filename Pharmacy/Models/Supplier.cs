using System.ComponentModel.DataAnnotations;

namespace Pharmacy.Models;

public class Supplier
{
    public int Id { get; set; }

    [Required(ErrorMessage = "Укажите название поставщика")]
    [StringLength(150)]
    [Display(Name = "Название")]
    public string Name { get; set; } = string.Empty;

    [StringLength(30)]
    [Display(Name = "Телефон")]
    public string? Phone { get; set; }

    [StringLength(200)]
    [Display(Name = "Адрес")]
    public string? Address { get; set; }

    public ICollection<Supply> Supplies { get; set; } = new List<Supply>();
}
