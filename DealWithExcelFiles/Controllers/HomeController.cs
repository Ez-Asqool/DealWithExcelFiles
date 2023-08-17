using DealWithExcelFiles.Data;
using DealWithExcelFiles.Models;
using DealWithExcelFiles.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Diagnostics;

namespace DealWithExcelFiles.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDBContext _dbContext;

        public HomeController(ILogger<HomeController> logger, ApplicationDBContext dBContext)
        {
            _logger = logger;
            _dbContext = dBContext;
        }

        public IActionResult Index()
        {
            return View();
        }

        //[HttpPost]
        //public async Task<IActionResult> Upload(FileUploadViewModel model)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            await model.ExcelFile.CopyToAsync(memoryStream);
        //            using (var package = new ExcelPackage(memoryStream))
        //            {
        //                var worksheet = package.Workbook.Worksheets.First(); // Assuming you have only one sheet

        //                // Process Excel data and insert into the database
        //                List<Product> products = new List<Product>();

        //                for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
        //                {
        //                    var sku = worksheet.Cells[row, 5].Text; 

        //                    // Skip rows with empty SKU (you can add more checks if needed)
        //                    if (string.IsNullOrWhiteSpace(sku))
        //                    {
        //                        continue;
        //                    }

        //                    var product = new Product();

        //                    int band;
        //                    if (int.TryParse(worksheet.Cells[row, 2].Text, out band))
        //                    {
        //                        product.Band = band;
        //                    }

        //                    product.CategoryCode = worksheet.Cells[row, 3].Text;
        //                    product.Manufacturer = worksheet.Cells[row, 4].Text;
        //                    product.PartSKU = sku;
        //                    product.ItemDescription = worksheet.Cells[row, 6].Text;
        //                    product.ListPrice = worksheet.Cells[row, 7].Text;
        //                    product.MinimumDiscount = worksheet.Cells[row, 8].Text;
        //                    product.DiscountedPrice = worksheet.Cells[row, 9].Text;

        //                    products.Add(product);
        //                }


        //                // Insert data into the database in batches (adjust batchSize as needed)
        //                int batchSize = 100;
        //                for (int i = 0; i < products.Count; i += batchSize)
        //                {
        //                    _dbContext.Products.AddRange(products.Skip(i).Take(batchSize));
        //                    await _dbContext.SaveChangesAsync();
        //                }
        //            }
        //        }

        //        return RedirectToAction("Success");
        //    }

        //    return View(model);
        //}


        //[HttpPost]
        //public async Task<IActionResult> Upload(FileUploadViewModel model)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            await model.ExcelFile.CopyToAsync(memoryStream);
        //            using (var package = new ExcelPackage(memoryStream))
        //            {
        //                var worksheets = package.Workbook.Worksheets;

        //                foreach (var worksheet in worksheets)
        //                {
        //                    // Process Excel data and insert into the database
        //                    List<Product> products = new List<Product>();

        //                    for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
        //                    {
        //                        var sku = worksheet.Cells[row, 5].Text; 

        //                        // Skip rows with empty SKU (you can add more checks if needed)
        //                        if (string.IsNullOrWhiteSpace(sku))
        //                        {
        //                            continue;
        //                        }

        //                        var product = new Product
        //                        {
        //                            Band = int.Parse(worksheet.Cells[row, 2].Text),
        //                            CategoryCode = worksheet.Cells[row, 3].Text,
        //                            Manufacturer = worksheet.Cells[row, 4].Text,
        //                            PartSKU = sku,
        //                            ItemDescription = worksheet.Cells[row, 6].Text,
        //                            ListPrice = worksheet.Cells[row, 7].Text,
        //                            MinimumDiscount = worksheet.Cells[row, 8].Text,
        //                            DiscountedPrice = worksheet.Cells[row, 9].Text
        //                        };

        //                        products.Add(product);
        //                    }

        //                    // Insert data into the database in batches (adjust batchSize as needed)
        //                    int batchSize = 100;
        //                    for (int i = 0; i < products.Count; i += batchSize)
        //                    {
        //                        _dbContext.Products.AddRange(products.Skip(i).Take(batchSize));
        //                        await _dbContext.SaveChangesAsync();
        //                    }
        //                }
        //            }
        //        }

        //        return RedirectToAction("Success");
        //    }

        //    return View(model);
        //}

        [HttpPost]
        public async Task<IActionResult> Upload(FileUploadViewModel model)
        {
            if (ModelState.IsValid)
            {
                using (var memoryStream = new MemoryStream())
                {
                    await model.ExcelFile.CopyToAsync(memoryStream);
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        var worksheets = package.Workbook.Worksheets;

                        foreach (var worksheet in worksheets)
                        {
                            for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var sku = worksheet.Cells[row, 5].Text;

                                // Skip rows with empty SKU (you can add more checks if needed)
                                if (string.IsNullOrWhiteSpace(sku))
                                {
                                    continue;
                                }

                                var existingProduct = _dbContext.Products.FirstOrDefault(p => p.PartSKU == sku);

                                if (existingProduct == null)
                                {
                                    // Insert new product if SKU doesn't exist
                                    var product = new Product
                                    {
                                        Band = int.Parse(worksheet.Cells[row, 2].Text),
                                        CategoryCode = worksheet.Cells[row, 3].Text,
                                        Manufacturer = worksheet.Cells[row, 4].Text,
                                        PartSKU = sku,
                                        ItemDescription = worksheet.Cells[row, 6].Text,
                                        ListPrice = worksheet.Cells[row, 7].Text,
                                        MinimumDiscount = worksheet.Cells[row, 8].Text,
                                        DiscountedPrice = worksheet.Cells[row, 9].Text
                                    };

                                    _dbContext.Products.Add(product);
                                }
                                else
                                {
                                    // Update existing product if SKU exists
                                    existingProduct.Band = int.Parse(worksheet.Cells[row, 2].Text);
                                    existingProduct.CategoryCode = worksheet.Cells[row, 3].Text;
                                    existingProduct.Manufacturer = worksheet.Cells[row, 4].Text;
                                    existingProduct.ItemDescription = worksheet.Cells[row, 6].Text;
                                    existingProduct.ListPrice = worksheet.Cells[row, 7].Text;
                                    existingProduct.MinimumDiscount = worksheet.Cells[row, 8].Text;
                                    existingProduct.DiscountedPrice = worksheet.Cells[row, 9].Text;

                                    _dbContext.Products.Update(existingProduct);
                                }
                            }

                            await _dbContext.SaveChangesAsync();
                        }
                    }
                }

                return RedirectToAction("Success");
            }

            return View(model);
        }



        [HttpGet]
        public IActionResult Success()
        {
            return View();
        }



        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}