using System.ComponentModel.DataAnnotations;

namespace DealWithExcelFiles.ViewModels
{
    public class FileUploadViewModel
    {
        [Required]
        [Display(Name = "Excel File")]
        public IFormFile ExcelFile { get; set; }
    }
}
