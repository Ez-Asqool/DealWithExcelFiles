using DealWithExcelFiles.Data;
using DealWithExcelFiles.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Drawing.Printing;
using System.Linq.Dynamic.Core;

namespace DealWithExcelFiles.Controllers
{
    public class ProductController : Controller
    {
        private readonly ApplicationDBContext _dBContext;

        public ProductController(ApplicationDBContext dBContext) 
        {
            _dBContext = dBContext;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult AllData() {
            var pageSize = int.Parse(Request.Form["length"]);
            var skip = int.Parse(Request.Form["start"]);

            string searchValue = Request.Form["search[value]"];

            var sortColumn = Request.Form[string.Concat("columns[", Request.Form["order[0][column]"], "][name]")];
            var sortColumnDirection = Request.Form["order[0][dir]"];


            IQueryable<Product> books = _dBContext.Products.AsQueryable();
            if (!string.IsNullOrEmpty(searchValue))
            {
                books = books.Where(x =>
                string.IsNullOrEmpty(searchValue) ? true :
                (x.CategoryCode.Contains(searchValue)) ||
                (x.Band.ToString().Contains(searchValue)) ||
                (x.Manufacturer.Contains(searchValue)) ||
                (x.ItemDescription.Contains(searchValue)) ||
                (x.ListPrice.Contains(searchValue)) ||
                (x.MinimumDiscount.Contains(searchValue)) ||
                (x.DiscountedPrice.Contains(searchValue)) ||
                (x.PartSKU.Contains(searchValue)));
            }

            if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
            {
                books = books.OrderBy(string.Concat(sortColumn, " ", sortColumnDirection));
            }

            var data = books.Skip(skip).Take(pageSize).ToList();

            var recordsTotal = books.Count();

            var jsonData = new
            {
                recordsFiltered = recordsTotal,
                recordsTotal,
                data
            };

            return Ok(jsonData);

        }





        [HttpPost]
        public IActionResult ExportToExcel(string searchValue)
        {
            IQueryable<Product> products = _dBContext.Products.AsQueryable();

            if (!string.IsNullOrEmpty(searchValue))
            {
                products = products.Where(x =>
                    x.CategoryCode.Contains(searchValue) ||
                    x.Band.ToString().Contains(searchValue) ||
                    x.Manufacturer.Contains(searchValue) ||
                    x.ItemDescription.Contains(searchValue) ||
                    x.ListPrice.Contains(searchValue) ||
                    x.MinimumDiscount.Contains(searchValue) ||
                    x.DiscountedPrice.Contains(searchValue) ||
                    x.PartSKU.Contains(searchValue));
            }

            var data = products.ToList();

            // Create a new Excel package
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Products");

                // Add column names in the first row
                string[] columnNames = { "Band #", "Category Code", "Manufacturer", "Part SKU", "Item Description", "List Price", "Minimum Discount", "Discounted Price" };
                for (int col = 0; col < columnNames.Length; col++)
                {
                    worksheet.Cells[1, col + 1].Value = columnNames[col];
                }

                // Populate the Excel sheet with data
                for (int row = 0; row < data.Count; row++)
                {
                    //worksheet.Cells[row + 2, 1].Value = row + 1; // Numbering
                    worksheet.Cells[row + 2, 1].Value = data[row].Band;
                    worksheet.Cells[row + 2, 2].Value = data[row].CategoryCode;
                    worksheet.Cells[row + 2, 3].Value = data[row].Manufacturer;
                    worksheet.Cells[row + 2, 4].Value = data[row].PartSKU;
                    worksheet.Cells[row + 2, 5].Value = data[row].ItemDescription;
                    worksheet.Cells[row + 2, 6].Value = data[row].ListPrice;
                    worksheet.Cells[row + 2, 7].Value = data[row].MinimumDiscount;
                    worksheet.Cells[row + 2, 8].Value = data[row].DiscountedPrice;
                }

                // Save the Excel package to a memory stream
                using (var memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    var excelBytes = memoryStream.ToArray();

                    // Save the Excel file to a temporary location
                    var tempFilePath = Path.Combine(Path.GetTempPath(), "Products.xlsx");
                    System.IO.File.WriteAllBytes(tempFilePath, excelBytes);

                    // Return the temporary file path to the client
                    return Json(new { fileName = "Products.xlsx" });
                }
            }
        }



        [HttpGet]
        public IActionResult DownloadExcel(string fileName)
        {
            var fileBytes = System.IO.File.ReadAllBytes(Path.Combine(Path.GetTempPath(), fileName));
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        //[HttpPost]
        //public IActionResult ExportToExcel(List<List<Product>> data)
        //{
        //    // Create a new Excel package
        //    using (var package = new ExcelPackage())
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        //        // Populate the Excel sheet with data
        //        for (int row = 0; row < data.Count; row++)
        //        {
        //            for (int col = 0; col < data[row].Count; col++)
        //            {
        //                worksheet.Cells[row + 1, col + 1].Value = data[row][col];
        //            }
        //        }

        //        // Save the Excel package to a memory stream
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            package.SaveAs(memoryStream);
        //            var excelBytes = memoryStream.ToArray();

        //            // Save the Excel file to a temporary location
        //            var tempFilePath = Path.Combine(Path.GetTempPath(), "temp_excel.xlsx");
        //            System.IO.File.WriteAllBytes(tempFilePath, excelBytes);

        //            // Return the temporary file path to the client
        //            return Json(new { fileName = "temp_excel.xlsx" });
        //        }
        //    }
        //}

        //[HttpGet]
        //public IActionResult DownloadExcel(string fileName)
        //{
        //    var fileBytes = System.IO.File.ReadAllBytes(Path.Combine(Path.GetTempPath(), fileName));
        //    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}

    }
}
