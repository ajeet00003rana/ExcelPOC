using ExcelPOC.Common;
using ExcelPOC.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using System.Net.Mime;

namespace ExcelPOC.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public HomeController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Index()
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

        [HttpGet]
        public async Task<IActionResult> ExportToExcel()
        {
            DataSet data = new DataSet();
            var table = SampleData.GetDataTable();
            data.Tables.Add(table.Copy());
            string fileName = "Excel_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string excelFilename = Path.Combine(_webHostEnvironment.WebRootPath, "Download", fileName);

            var bytes = await ExcelService.CreateExcel(data, excelFilename, (DataSet ds, string file) =>
                {
                    CommonFunction.ExportDataSetWithInBuiltFunction(ds, file);
                    return Task.CompletedTask;
                });
            return File(bytes, MediaTypeNames.Application.Octet, fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportToExcelDefault()
        {
            DataSet data = new DataSet();
            var table = SampleData.GetDataTable();
            data.Tables.Add(table.Copy());
            string fileName = "Excel_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string excelFilename = Path.Combine(_webHostEnvironment.WebRootPath, "Download", fileName);

            var bytes = await ExcelService.CreateExcel(data, excelFilename, ExportToExcelDefault);
            return File(bytes, MediaTypeNames.Application.Octet, fileName);
        }

        private Task ExportToExcelDefault(DataSet ds, string file)
        {
            CommonFunction.ExportDataSetWithCustomLogic(ds, file);
            return Task.CompletedTask;
        }
    }
}