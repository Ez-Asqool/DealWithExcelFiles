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
                        //int totalRowCount = 0;

                        var worksheets = package.Workbook.Worksheets;

                        foreach (var worksheet in worksheets)
                        {
                            var batch = new List<Product>();
                            int countEmptyRows = 0;

                            for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var sku = worksheet.Cells[row, 5].Text;

                                // Skip rows with empty SKU (you can add more checks if needed)
                                if (string.IsNullOrWhiteSpace(sku))
                                {
                                    if (countEmptyRows == 5)
                                    {
                                        break;
                                    }

                                    countEmptyRows++;
                                    continue;

                                }

                                var existingProduct = _dbContext.Products.FirstOrDefault(p => p.PartSKU == sku);

                                if (existingProduct == null)
                                {
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

                                    batch.Add(product);
                                }
                                else
                                {
                                    existingProduct.Band = int.Parse(worksheet.Cells[row, 2].Text);
                                    existingProduct.CategoryCode = worksheet.Cells[row, 3].Text;
                                    existingProduct.Manufacturer = worksheet.Cells[row, 4].Text;
                                    existingProduct.ItemDescription = worksheet.Cells[row, 6].Text;
                                    existingProduct.ListPrice = worksheet.Cells[row, 7].Text;
                                    existingProduct.MinimumDiscount = worksheet.Cells[row, 8].Text;
                                    existingProduct.DiscountedPrice = worksheet.Cells[row, 9].Text;

                                    _dbContext.Products.Update(existingProduct);
                                }

                                if (batch.Count >= 20000) // Adjust the batch size as needed
                                {
                                    _dbContext.Products.AddRange(batch);
                                    batch.Clear();
                                    await _dbContext.SaveChangesAsync();

                                    Thread.Sleep(2000);
                                }


                            }

                            if (batch.Any())
                            {
                                _dbContext.Products.AddRange(batch);
                                await _dbContext.SaveChangesAsync();
                            }
                        }
                    }
                }
                return RedirectToAction("Success");
            }

            return RedirectToAction("Index");
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