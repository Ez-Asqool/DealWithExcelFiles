using System.ComponentModel.DataAnnotations;

namespace DealWithExcelFiles.Models
{
    public class Product
    {
        public int Id { get; set; }


        [Required]
        public int Band { get; set; }

        
        [Required]
        public string CategoryCode { get; set; }

        
        [Required]
        public string Manufacturer { get; set; }

        
        [Required]
        public string PartSKU { get; set; }


        [Required]
        public string ItemDescription { get; set; }

        [Required]
        public string ListPrice { get; set; }


        [Required]
        public string MinimumDiscount { get; set; }


        [Required]
        public string DiscountedPrice { get; set; }


    }
}
