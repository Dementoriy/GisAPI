using Dadata;
using GisAPI.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;

namespace GisAPI.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private string filePathExcell = Path.Combine(Directory.GetCurrentDirectory(), $"Report" + ".xlsx");
        private string fullPath = Path.Combine(Directory.GetCurrentDirectory(), $"file.xlsx");
        ISheet sheet;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Map()
        {
            XSSFWorkbook xssfwb;
            using (FileStream fileStream = new FileStream(fullPath, FileMode.Open))
            {
                xssfwb = new XSSFWorkbook(fileStream);
            }
            ISheet excelSheet = xssfwb.GetSheetAt(0);

            filePathExcell = FillPathExcell(excelSheet, GeoCoordinates.coordinates);
            System.IO.File.Delete(fullPath);
            return PhysicalFile(filePathExcell, "application/xlsx", "Report.xlsx");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IActionResult Index(Addresses addresses)
        {
            var file = Request.Form.Files.First();
            ViewBag.addresses = GetAddresses(file);
            return View("Map");
        }

        public double[][] GetAddresses(IFormFile file)
        {
            fullPath = Path.Combine(Directory.GetCurrentDirectory(), $"file.xlsx");
            XSSFWorkbook xssfwb;
            using (FileStream fileStream = new FileStream(fullPath, FileMode.Create))
            {
                file.CopyTo(fileStream);
            }
            using (FileStream fileStream = new FileStream(fullPath, FileMode.Open))
            {
                xssfwb = new XSSFWorkbook(fileStream);
            }
            var daDatatoken = "a8736e9ed0f081d84ba27c45c4fc86ddac04183a";
            var api = new SuggestClient(daDatatoken);

            sheet = xssfwb.GetSheetAt(0);

            GeoCoordinates.coordinates = new double[sheet.LastRowNum][];
            //coordinates = new double[sheet.LastRowNum][];

            for (int row = 6; row <= sheet.LastRowNum; row++)
            {
                string address = "";

                if (sheet.GetRow(row) != null)
                {
                    address += sheet.GetRow(row).GetCell(12).ToString();
                }
                
                var result = api.SuggestAddress(address);

                if (result.suggestions.Count == 0)
                {
                    break;
                }
                GeoCoordinates.coordinates[row-6] = new double[2];
                GeoCoordinates.coordinates[row-6][0] = Convert.ToDouble(result.suggestions[0].data.geo_lon.Replace(".", ","));
                GeoCoordinates.coordinates[row-6][1] = Convert.ToDouble(result.suggestions[0].data.geo_lat.Replace(".", ","));


            }
            GeoCoordinates.coordinates = GeoCoordinates.coordinates.Where(c => c != null).ToArray();
            return GeoCoordinates.coordinates;
        }

        public static string FillPathExcell(ISheet excelSheet, double[][] coordinates)
        {
            string path = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(path, $"Report" + ".xlsx");

            if(System.IO.File.Exists(filePath))
                System.IO.File.Delete(filePath);

            var file = System.IO.File.Create(filePath);
            var template = new MemoryStream(Properties.Resources.shablon, true);
            var workbook = new XSSFWorkbook(template);
            var sheet = workbook.GetSheetAt(0);

            sheet.ShiftRows(1, 1 + excelSheet.LastRowNum, excelSheet.LastRowNum, true, true);
            int row = 1;

            for(int i = 6; i <= excelSheet.LastRowNum; i++)
            {
                var rowInsert = sheet.CreateRow(row);
                
                rowInsert.CreateCell(0).SetCellValue(excelSheet.GetRow(i).GetCell(10).ToString());
                rowInsert.GetCell(0).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                rowInsert.CreateCell(1).SetCellValue(excelSheet.GetRow(i).GetCell(12).ToString());
                rowInsert.GetCell(1).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                rowInsert.CreateCell(2).SetCellValue(excelSheet.GetRow(i).GetCell(6).ToString());
                rowInsert.GetCell(2).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                rowInsert.CreateCell(3).SetCellValue(excelSheet.GetRow(i).GetCell(9).ToString());
                rowInsert.GetCell(3).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                rowInsert.CreateCell(4).SetCellValue(excelSheet.GetRow(i).GetCell(11).ToString());
                rowInsert.GetCell(4).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                row++;
            }

            for (int i = 0; i < coordinates.Length; i++)
            {
                var rowInsert = sheet.GetRow(i + 1);

                rowInsert.CreateCell(5).SetCellValue(coordinates[i][1]);
                rowInsert.GetCell(5).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
                rowInsert.CreateCell(6).SetCellValue(coordinates[i][0]);
                rowInsert.GetCell(6).CellStyle = sheet.GetRow(0).GetCell(0).CellStyle;
            }

            workbook.Write(file);
            workbook.Close();
            template.Close();
            file.Close();

            return (filePath);
        }

    }
}