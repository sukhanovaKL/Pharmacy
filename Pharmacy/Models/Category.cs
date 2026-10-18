using System.ComponentModel.DataAnnotations;

namespace Pharmacy.Models;

public class Category
{
    public int Id { get; set; }

    [Required(ErrorMessage = "Укажите название категории")]
    [StringLength(100)]
    [Display(Name = "Название")]
    public string Name { get; set; } = string.Empty;

    public ICollection<Medicine> Medicines { get; set; } = new List<Medicine>();
}
