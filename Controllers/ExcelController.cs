using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;
using MySqlConnector;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IConfiguration _config;

        public ExcelController(IConfiguration config)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _config = config;
        }

        [HttpPost("upload")]
        public IActionResult UploadExcel([FromForm] IFormFile file, [FromForm] int type)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Dosya yüklenmedi");

            try
            {
                Console.WriteLine($"[Request] UploadExcel çağrıldı. Zaman: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                Console.WriteLine($"Dosya alındı: {file.FileName}, Boyut: {file.Length} bytes");

                using var stream = new MemoryStream();
                file.CopyTo(stream);
                stream.Position = 0;

                var workbook = new Workbook();
                workbook.Open(stream);

                var karne = workbook.Worksheets["PARAKENDECİLİK(KARNE)"];
                var bayi = workbook.Worksheets["bayi-personel"];
                if (karne == null || bayi == null)
                    return BadRequest("Gerekli sayfalar bulunamadı.");

                var usedRange = bayi.UsedRange;
                int lastRow = usedRange.Row + usedRange.RowCount - 1;

                var dValues = bayi.Range[$"D2:D{lastRow}"].Value as object[,];
                var cValues = bayi.Range[$"C2:C{lastRow}"].Value as object[,];

                var pairs = new List<(string d, string c)>();
                for (int i = 0; i < dValues.GetLength(0); i++)
                {
                    var d = dValues[i, 0]?.ToString()?.Trim();
                    var c = cValues[i, 0]?.ToString()?.Trim();

                    if (!string.IsNullOrEmpty(d) && !string.IsNullOrEmpty(c))
                        pairs.Add((d, c));
                }

                Console.WriteLine($"İşlenecek kayıt sayısı: {pairs.Count}");

                var resultList = new List<object>();
                var dbItems = new List<(string id, string base64)>();

                var options = new ImageSaveOptions
                {
                    Resolution = 150,
                    ShowGridlines = false,
                    ShowColumnHeadings = false,
                    ShowRowHeadings = false
                };

                int processedCount = 0;

                foreach (var (dVal, cVal) in pairs)
                {
                    karne.Range["C2"].Value = dVal;
                    karne.Range["B4"].Value = cVal;

                    workbook.Calculate();

                    using var imageStream = new MemoryStream();
                    karne.Range["A1:AE55"].ToImage(imageStream, ImageType.PNG, options);

                    var bytes = imageStream.ToArray();
                    string base64 = Convert.ToBase64String(bytes);
                    string id = ExtractCode(cVal);

                    resultList.Add(new { id });
                    dbItems.Add((id, base64));

                    processedCount++;
                    Console.WriteLine($"{processedCount}. kayıt işlendi. ID: {id}");
                }

                Console.WriteLine($"Toplam {processedCount} kayıt işlendi ve base64'e dönüştürüldü.");

                SaveToDatabase(dbItems, type);

                return Ok(resultList);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Sunucu hatası: {ex.Message}");
            }
        }

        private void SaveToDatabase(List<(string id, string base64)> items, int type)
        {
            string connStr = _config.GetConnectionString("MySql")!;
            using var conn = new MySqlConnection(connStr);
            conn.Open();

            using (var createCmd = conn.CreateCommand())
            {
                createCmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS tbl_reports (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        branch_code VARCHAR(255),
                        img LONGTEXT,
                        upload_date DATETIME,
                        type INT,
                        INDEX idx_branch_code (branch_code)
                    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci";
                createCmd.ExecuteNonQuery();
            }

            using (var deleteCmd = conn.CreateCommand())
            {
                deleteCmd.CommandText = "DELETE FROM tbl_reports WHERE type = @type";
                deleteCmd.Parameters.AddWithValue("@type", type);
                int deleted = deleteCmd.ExecuteNonQuery();
                Console.WriteLine($"MySQL'den {deleted} eski kayıt silindi. Type: {type}");
            }

            foreach (var (branchCode, base64) in items)
            {
                using var insertCmd = conn.CreateCommand();
                insertCmd.CommandText = "INSERT INTO tbl_reports (branch_code, img, upload_date, type) VALUES (@branch_code, @img, NOW(), @type)";
                insertCmd.Parameters.AddWithValue("@branch_code", branchCode);
                insertCmd.Parameters.AddWithValue("@img", base64);
                insertCmd.Parameters.AddWithValue("@type", type);
                insertCmd.ExecuteNonQuery();
            }

            Console.WriteLine($"MySQL'e {items.Count} kayıt yazıldı. Type: {type}");
        }

        private string ExtractCode(string value)
        {
            var start = value.IndexOf("00000.");
            if (start == -1) return Sanitize(value);

            var sub = value[start..];
            var end = sub.IndexOf('_');
            return Sanitize(end != -1 ? sub[..end] : sub);
        }

        private string Sanitize(string input)
            => string.Concat(input.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
    }
}
