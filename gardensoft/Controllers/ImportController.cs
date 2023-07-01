using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

using gardensoft.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Newtonsoft.Json;
using System.Text;

namespace gardensoft.Controllers
{
    public class ImportController : Controller
    {
        private readonly ILogger<ImportController> _logger;
        List<KhachHang> khachhangList = new List<KhachHang>();


        public ImportController(ILogger<ImportController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        //[HttpPost]
        public IActionResult ImportExcel([FromForm] IFormFile excelFile, int page = 1, int pageSize = 10)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return RedirectToAction("Error");
            }

            using (ExcelPackage package = new ExcelPackage(excelFile.OpenReadStream()))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    KhachHang khachhang = new KhachHang();

                    khachhang.MaID = worksheet.Cells[row, 1].Value?.ToString();
                    khachhang.Ten = worksheet.Cells[row, 2].Value?.ToString();
                    khachhang.NgaySinh = Convert.ToDateTime(worksheet.Cells[row, 3].Value);
                    khachhang.DiaChi = worksheet.Cells[row, 4].Value?.ToString();
                    khachhang.PassPort = worksheet.Cells[row, 5].Value?.ToString();
                    khachhang.NgayCap = Convert.ToDateTime(worksheet.Cells[row, 6].Value);
                    khachhang.DienThoai = worksheet.Cells[row, 7].Value?.ToString();
                    khachhang.DiDong = worksheet.Cells[row, 8].Value?.ToString();
                    khachhang.Fax = worksheet.Cells[row, 9].Value?.ToString();
                    khachhang.Email = worksheet.Cells[row, 10].Value?.ToString();
                    khachhang.TaiKhoanNH = worksheet.Cells[row, 11].Value?.ToString();
                    khachhang.TenNH = worksheet.Cells[row, 12].Value?.ToString();
                    khachhang.LoaiKH = worksheet.Cells[row, 13].Value?.ToString();
                    khachhang.HanTT = worksheet.Cells[row, 14].Value?.ToString();

                    khachhangList.Add(khachhang);
                }
            }
            HttpContext.Session.Set("KhachhangList", ConvertToByteArray(khachhangList));
            int totalItems = khachhangList.Count;
            int totalPages = (int)Math.Ceiling(totalItems / (double)pageSize);

            List<KhachHang> pagedKhachHangList = khachhangList
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            ViewData["KhachhangList"] = pagedKhachHangList;
            ViewBag.TotalPages = totalPages;
            ViewBag.CurrentPage = page;

            return View("Index");
        }

        [HttpPost]
        public IActionResult insertDataFromExcel()
        {
            var khachhangList = HttpContext.Session.Get("KhachhangList");
            List<KhachHang> khachhangList1 = ConvertFromByteArray<List<KhachHang>>(khachhangList);

            string sqlConnect = "Data Source=DESKTOP-101QR58;Initial Catalog=QLKH;Integrated Security=True";
            int count = 0;
            using (SqlConnection connection = new SqlConnection(sqlConnect))
            {
                connection.Open();
                foreach (var kh in khachhangList1)
                {
                    if (
                            (kh.MaID != null && !kh.MaID.StartsWith("#")) &&
                            (kh.Ten != null && !kh.Ten.StartsWith("#")) &&
                            (kh.NgaySinh.ToString() != null && !kh.NgaySinh.ToString().StartsWith("#")) &&
                            (kh.DiaChi != null && !kh.DiaChi.StartsWith("#")) &&
                            (kh.PassPort != null && !kh.PassPort.StartsWith("#")) &&
                            (kh.NgayCap.ToString() != null && !kh.NgayCap.ToString().StartsWith("#")) &&
                            (kh.DienThoai != null && !kh.DienThoai.StartsWith("#")) &&
                            (kh.DiDong != null && !kh.DiDong.StartsWith("#")) &&
                            (kh.Fax != null && !kh.Fax.StartsWith("#")) &&
                            (kh.Email != null && !kh.Email.StartsWith("#")) &&
                            (kh.TaiKhoanNH != null && !kh.TaiKhoanNH.StartsWith("#")) &&
                            (kh.TenNH != null && !kh.TenNH.StartsWith("#")) &&
                            (kh.LoaiKH != null && !kh.LoaiKH.StartsWith("#")) &&
                            (kh.HanTT != null && !kh.HanTT.StartsWith("#"))
                        )
                    {
                        using (var command = new SqlCommand("InsertKhachHang", connection))
                        {
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.AddWithValue("@MaID", kh.MaID);
                            command.Parameters.AddWithValue("@Ten", kh.Ten);
                            command.Parameters.AddWithValue("@NgaySinh", kh.NgaySinh);
                            command.Parameters.AddWithValue("@DiaChi", kh.DiaChi);
                            command.Parameters.AddWithValue("@PassPort", kh.PassPort);
                            command.Parameters.AddWithValue("@NgayCap", kh.NgayCap);
                            command.Parameters.AddWithValue("@DienThoai", kh.DienThoai);
                            command.Parameters.AddWithValue("@DiDong", kh.DiDong);
                            command.Parameters.AddWithValue("@Fax", kh.Fax);
                            command.Parameters.AddWithValue("@Email", kh.Email);
                            command.Parameters.AddWithValue("@TaiKhoanNH", kh.TaiKhoanNH);
                            command.Parameters.AddWithValue("@TenNH", kh.TenNH);
                            command.Parameters.AddWithValue("@LoaiKH", kh.LoaiKH);
                            command.Parameters.AddWithValue("@HanTT", kh.HanTT);

                            SqlParameter successParam = new SqlParameter("@Success", SqlDbType.Bit)
                            {
                                Direction = ParameterDirection.Output
                            };
                            command.Parameters.Add(successParam);

                            command.ExecuteNonQuery();
                            bool success = (bool)successParam.Value;
                            if (success)
                            {
                                count++;
                            }
                        }
                    }
                    
                }
                
                string resultMessage = "Đã thêm thành công " + count + " record";

                TempData["result"] = resultMessage;
                return RedirectToAction("Index");
            }
        }

        private byte[] ConvertToByteArray<T>(T obj)
        {
            string json = JsonConvert.SerializeObject(obj);
            return Encoding.UTF8.GetBytes(json);
        }

        private T ConvertFromByteArray<T>(byte[] data)
        {
            string json = Encoding.UTF8.GetString(data);
            return JsonConvert.DeserializeObject<T>(json);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
