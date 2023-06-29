using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;

using gardensoft.Models;

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace gardensoft.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult Input()
        {
            List<KhachHang> khachhangList = new List<KhachHang>();

            string sqlConnect = "Data Source=DESKTOP-AH3TGNG;Initial Catalog=QLKH;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(sqlConnect))
            {
                connection.Open();
                string sql = "SELECT * FROM KHACHHANG";
                using (SqlCommand comand = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = comand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            KhachHang kh = new KhachHang();

                            kh.MaID = reader.GetString(0);
                            kh.Ten = reader.GetString(1);
                            kh.NgaySinh = reader.GetDateTime(2);
                            kh.DiaChi = reader.GetString(3);
                            kh.PassPort = reader.GetString(4);
                            kh.NgayCap = reader.GetDateTime(5);
                            kh.DienThoai = reader.GetString(6);
                            kh.DiDong = reader.GetString(7);
                            kh.Fax = reader.GetString(8);
                            kh.Email = reader.GetString(9);
                            kh.TaiKhoanNH = reader.GetString(10);
                            kh.TenNH = reader.GetString(11);
                            kh.LoaiKH = reader.GetString(12);
                            kh.HanTT = reader.GetString(13);

                            khachhangList.Add(kh);
                        }
                    }
                }
                connection.Close();
            }
            ViewData["KhachhangList"] = khachhangList;
            return View("Input");
        }

        [HttpPost]
        public IActionResult PostInput()
        {
            try
            {
                KhachHang kh = new KhachHang();

                kh.MaID = HttpContext.Request.Form["maID"];
                kh.Ten = HttpContext.Request.Form["ten"];
                kh.NgaySinh = DateTime.ParseExact(HttpContext.Request.Form["ngaySinh"] + "", "yyyy-MM-dd", CultureInfo.InvariantCulture);
                kh.DiaChi = HttpContext.Request.Form["diaChi"];
                kh.PassPort = HttpContext.Request.Form["passPort"];
                kh.NgayCap = DateTime.ParseExact(HttpContext.Request.Form["ngayCap"] + "", "yyyy-MM-dd", CultureInfo.InvariantCulture);
                kh.DienThoai = HttpContext.Request.Form["dienThoai"];
                kh.DiDong = HttpContext.Request.Form["diDong"];
                kh.Fax = HttpContext.Request.Form["fax"];
                kh.Email = HttpContext.Request.Form["email"];
                kh.TaiKhoanNH = HttpContext.Request.Form["taiKhoanNH"];
                kh.TenNH = HttpContext.Request.Form["tenNH"];
                kh.LoaiKH = HttpContext.Request.Form["loaiKH"];
                kh.HanTT = HttpContext.Request.Form["hanTT"];


                string sqlConnect = "Data Source=DESKTOP-AH3TGNG;Initial Catalog=QLKH;Integrated Security=True";
                using (SqlConnection connection = new SqlConnection(sqlConnect))
                {
                    connection.Open();
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
                        string resultMessage = success ? "Thêm thành công!" : "Đã tồn tại mã ID!!!";

                        TempData["result"] = resultMessage;
                        return RedirectToAction("Input");
                    }
                }
            }
            catch (Exception ex)
            {
                string result = ex.Message;
                TempData["result"] = result;
                return RedirectToAction("Input");
            }
        }

        public IActionResult ExportFile()
        {
            string sqlConnect = "Data Source=DESKTOP-AH3TGNG;Initial Catalog=QLKH;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(sqlConnect))
            {
                connection.Open();

                string sqlQuery = "SELECT * FROM KHACHHANG";

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        ToExcel(reader, "QLKH");
                    }
                }
            }
            TempData["result"] = "Export file thành công!!!";

            return RedirectToAction("Input");
        }

        private void ToExcel(SqlDataReader reader, string baseFileName)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                string fileName = GetUniqueFileName(baseFileName);
                workbook = excel.Workbooks.Add(Type.Missing);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Quản lý khách hàng";

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    worksheet.Cells[1, i + 1] = reader.GetName(i);
                }

                int row = 2;

                while (reader.Read())
                {
                    for (int col = 0; col < reader.FieldCount; col++)
                    {
                        worksheet.Cells[row, col + 1] = "'" + reader.GetValue(col).ToString();
                    }
                    row++;
                }

                workbook.SaveAs(fileName);
                workbook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
        }

        private string GetUniqueFileName(string baseFileName)
        {
            string fileExtension = Path.GetExtension(baseFileName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(baseFileName);

            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string uniqueFileName = fileNameWithoutExtension;
            int count = 1;

            while (System.IO.File.Exists(Path.Combine(path, uniqueFileName + fileExtension + ".xlsx")))
            {
                uniqueFileName = $"{fileNameWithoutExtension} ({count})";
                count++;
            }

            return Path.Combine(path, uniqueFileName + fileExtension + ".xlsx");
        }


        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}