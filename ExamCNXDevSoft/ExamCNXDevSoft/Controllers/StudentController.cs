using Microsoft.AspNetCore.Mvc;
using NuGet.Protocol;
using System.Net;
using ExamCNXDevSoft.Models;
using ExamCNXDevSoft.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using Microsoft.CodeAnalysis.Elfie.Diagnostics;

namespace ExamCNXDevSoft.Controllers
{
    public class StudentController : Controller
    {
        private readonly ApplicationDBContext _db;
        public StudentController(ApplicationDBContext db)
        {
            _db = db;
        }

        public IActionResult Index()
        {
            IEnumerable<Student> allStudent = _db.Students;
            return View(allStudent);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(Student obj)
        {
            if (ModelState.IsValid)
            {
                _db.Students.Add(obj);
                _db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(obj);
        }

        public IActionResult Edit(int? id)
        {
            if (id == null || id == 0)
            {
                return NotFound();

            }
            var obj = _db.Students.Find(id);
            if (obj == null)
            {
                return NotFound();

            }
            return View(obj);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(Student obj)
        {
            if (ModelState.IsValid)
            {
                _db.Students.Update(obj);
                _db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(obj);
        }

        public IActionResult Delete(int id)
        {
            if (id == null || id == 0)
            {
                return NotFound();

            }
            var obj = _db.Students.Find(id);
            if (obj == null)
            {
                return NotFound();

            }
            _db.Students.Remove(obj);
            _db.SaveChanges();
            return RedirectToAction("Index");
        }

        public IActionResult ExportExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                //sheet name
                var ws = workbook.Worksheets.Add("Report");


                //Header 
                ws.Cell(1, 1).Value = "รหัสนักเรียน";
                ws.Cell(1, 2).Value = "ชื่อนักเรียน";
                ws.Cell(1, 3).Value = "คะแนนสอบ";



                // Connect To SQL

                System.Data.DataTable dt = new System.Data.DataTable();
                SqlConnection con = new SqlConnection("Server = localhost\\SQLEXPRESS01; Database = studentDB; Trusted_Connection = True; TrustServerCertificate = True");
                DateTime dateTime = DateTime.Today.AddDays(-1);
                string datetime = dateTime.ToString();
                string singleQote = "'";
    
                SqlDataAdapter ad = new SqlDataAdapter("select * from Students", con);
                ad.Fill(dt);
                int i = 2;

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    int u = 1;
                    foreach (var item in row.ItemArray)
                    {

                        ws.Cell(i, u).Value = item.ToString();
                        u++;
                    }



                    i++;

                }



                //using (var stream = new MemoryStream())
                //{
                //    var date = DateTime.Today.AddDays(-1).ToString("dd MMMM yyyy");
                //    string filename = "ข้อมูลนักเรียน" + date + ".xlsx";

                //    string combinePath = System.IO.Directory.CreateDirectory("ExportFiles/") + filename;
                //    string pathSave = Path.GetFullPath(combinePath);
                //    workbook.SaveAs(pathSave); ;
                //    var content = stream.ToArray();

                //    return File
                //        (
                //        content,
                //        "application/vnd.ms-excel.sheet.macroEnabled.12",
                //        "ข้อมูลนักเรียน" + date + ".xlsx"
                //        );
                //}
                using (var steam  = new MemoryStream())
                {
                    workbook.SaveAs(steam);
                    var content = steam.ToArray();
                    var date = DateTime.Today.AddDays(-1).ToString("dd MMMM yyyy");
                    return File
                    (
                        content,
                       "application/vnd.ms-excel.sheet.macroEnabled.12",
                      "ข้อมูลนักเรียน" + date + ".xlsx"
                    );
                }
            }
        }

        public IActionResult ExternalApi()
        {

            using (var client = new HttpClient())
            {
                var endpoint = new Uri("https://jsonplaceholder.typicode.com/posts");
                var result = client.GetAsync(endpoint).Result;
                var json = result.Content.ReadAsStringAsync().Result;

                return Json(json);
            }







        }
    }
}
    

