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

            sheet = xssfwb.GetSheetAt(0);

            GeoCoordinates.coordinates = new double[sheet.LastRowNum][];

            var rows = Enumerable.Range(6, sheet.LastRowNum).Select(sheet.GetRow).Where(r => r != null).ToList();
            var tasks = new Task<double[][]>[2];

            tasks[0] = new Task<double[][]>(() => GetTask(rows.Take(new Range(0, rows.Count / 2))));
            tasks[1] = new Task<double[][]>(() => GetTask(rows.Take(new Range(rows.Count / 2, rows.Count))));

            tasks[0].Start();
            tasks[1].Start();
            GeoCoordinates.coordinates = tasks[0].Result.Concat(tasks[1].Result).ToArray();

            GeoCoordinates.coordinates = GeoCoordinates.coordinates.Where(c => c != null).ToArray();
            return GeoCoordinates.coordinates;
        }

        private double[][] GetTask(IEnumerable<IRow> rows)
        {
            var tmp = rows.ToArray();
            var res = new double[tmp.Length][];

            var daDatatoken = "a8736e9ed0f081d84ba27c45c4fc86ddac04183a";
            var api = new SuggestClient(daDatatoken);

            for (int i = 0; i < tmp.Length; i++)
            {
                string address = "";

                address += tmp[i].GetCell(12).ToString();

                var result = api.SuggestAddress(address);

                if (result.suggestions.Count == 0)
                {
                    break;
                }
                res[i] = new double[2];
                res[i][0] = Convert.ToDouble(result.suggestions[0].data.geo_lon.Replace(".", ","));
                res[i][1] = Convert.ToDouble(result.suggestions[0].data.geo_lat.Replace(".", ","));

            }
            return res;
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