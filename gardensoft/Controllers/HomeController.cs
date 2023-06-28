using System.Data.SqlClient;
using System.Diagnostics;

using gardensoft.Models;

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

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
            return View("Input", khachhangList);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}