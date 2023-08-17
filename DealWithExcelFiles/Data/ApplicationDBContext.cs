using DealWithExcelFiles.Models;
using Microsoft.EntityFrameworkCore;

namespace DealWithExcelFiles.Data
{
    public class ApplicationDBContext : DbContext
    {
        public ApplicationDBContext(DbContextOptions options) : base(options) 
        {
            
        }


        public DbSet<Product> Products { get; set; }

    }
}
