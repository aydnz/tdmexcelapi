using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;
using MySqlConnector;
using System.Diagnostics;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IConfiguration _config;
        private static readonly object _logLock = new object();
        private static readonly string _logFilePath = Path.Combine(AppContext.BaseDirectory, "upload_debug.log");

        public ExcelController(IConfiguration config)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _config = config;
        }

        // Postman'den GET /api/excel/log ile, upload isteği hâlâ devam ederken/timeout'a düşse bile
        // en son nereye kadar geldiğini görmek için ayrı bir istek olarak çağırılabilir.
        [HttpGet("log")]
        public IActionResult GetLog([FromQuery] int lines = 300)
        {
            try
            {
                if (!System.IO.File.Exists(_logFilePath))
                    return Content("Log dosyası henüz oluşmadı.", "text/plain; charset=utf-8");

                string[] allLines;
                lock (_logLock)
                {
                    allLines = System.IO.File.ReadAllLines(_logFilePath);
                }
                var tail = allLines.Skip(Math.Max(0, allLines.Length - lines));
                return Content(string.Join("\n", tail), "text/plain; charset=utf-8");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Log okunamadı: {ex.Message}");
            }
        }

        [HttpGet("log/clear")]
        public IActionResult ClearLog()
        {
            try
            {
                lock (_logLock)
                {
                    System.IO.File.WriteAllText(_logFilePath, string.Empty);
                }
                return Content("Log temizlendi.", "text/plain; charset=utf-8");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Log temizlenemedi: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            var line = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss.fff} - {message}";
            Console.WriteLine(line);
            try
            {
                lock (_logLock)
                {
                    System.IO.File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                // log yazılamazsa isteği etkilemesin
            }
        }

        [HttpPost("upload")]
        public IActionResult UploadExcel([FromForm] IFormFile file, [FromForm] int type)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Dosya yüklenmedi");

            var swTotal = Stopwatch.StartNew();

            try
            {
                Log($"[Request] UploadExcel çağrıldı. Dosya: {file.FileName}, Boyut: {file.Length} bytes, Type: {type}");

                var swParse = Stopwatch.StartNew();

                using var stream = new MemoryStream();
                file.CopyTo(stream);
                stream.Position = 0;

                var workbook = new Workbook();
                workbook.Open(stream);

                var karne = workbook.Worksheets["PARAKENDECİLİK(KARNE)"];
                var bayi = workbook.Worksheets["bayi-personel"];
                if (karne == null || bayi == null)
                {
                    Log("[Hata] Gerekli sayfalar bulunamadı.");
                    return BadRequest("Gerekli sayfalar bulunamadı.");
                }

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

                swParse.Stop();
                Log($"Parse tamamlandı ({swParse.ElapsedMilliseconds}ms). İşlenecek kayıt sayısı: {pairs.Count}");

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
                var swRender = Stopwatch.StartNew();
                Log("Render döngüsü başladı.");

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
                    if (processedCount % 10 == 0 || processedCount == pairs.Count)
                        Log($"{processedCount}/{pairs.Count} kayıt render edildi. Son ID: {id} (geçen süre: {swRender.ElapsedMilliseconds}ms)");
                }

                swRender.Stop();
                Log($"Render döngüsü bitti ({swRender.ElapsedMilliseconds}ms). Toplam {processedCount} kayıt işlendi.");

                Log("SaveToDatabase başlıyor.");
                var dbTimings = SaveToDatabase(dbItems, type);
                Log("SaveToDatabase bitti.");

                swTotal.Stop();

                Log($"[Timing] parse={swParse.ElapsedMilliseconds}ms render={swRender.ElapsedMilliseconds}ms " +
                    $"dbConnect={dbTimings.connectMs}ms dbCreate={dbTimings.createMs}ms dbDelete={dbTimings.deleteMs}ms " +
                    $"dbInsert={dbTimings.insertMs}ms dbCommit={dbTimings.commitMs}ms total={swTotal.ElapsedMilliseconds}ms");

                return Ok(resultList);
            }
            catch (Exception ex)
            {
                swTotal.Stop();
                Log($"[Hata] {ex.Message} (total={swTotal.ElapsedMilliseconds}ms)\n{ex.StackTrace}");
                return StatusCode(500, $"Sunucu hatası: {ex.Message}");
            }
        }

        private (long connectMs, long createMs, long deleteMs, long insertMs, long commitMs) SaveToDatabase(List<(string id, string base64)> items, int type)
        {
            string connStr = _config.GetConnectionString("MySql")!;
            using var conn = new MySqlConnection(connStr);

            Log("DB: bağlantı açılıyor (conn.Open)...");
            var swConnect = Stopwatch.StartNew();
            conn.Open();
            swConnect.Stop();
            Log($"DB: bağlantı açıldı ({swConnect.ElapsedMilliseconds}ms).");

            const int dbCommandTimeoutSeconds = 180;

            var swCreate = Stopwatch.StartNew();
            using (var createCmd = conn.CreateCommand())
            {
                createCmd.CommandTimeout = dbCommandTimeoutSeconds;
                createCmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS tbl_reports (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        branch_code VARCHAR(255),
                        img LONGTEXT,
                        upload_date DATETIME,
                        type INT,
                        INDEX idx_branch_code (branch_code),
                        INDEX idx_type (type)
                    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci";
                createCmd.ExecuteNonQuery();
            }
            swCreate.Stop();
            Log($"DB: create table kontrolü tamam ({swCreate.ElapsedMilliseconds}ms).");

            // Tablo daha önce (idx_type olmadan) oluşturulmuş olabilir - varsa dokunma, yoksa ekle.
            using (var indexCmd = conn.CreateCommand())
            {
                indexCmd.CommandTimeout = dbCommandTimeoutSeconds;
                indexCmd.CommandText = @"
                    SELECT COUNT(1) FROM information_schema.statistics
                    WHERE table_schema = DATABASE() AND table_name = 'tbl_reports' AND index_name = 'idx_type'";
                var exists = Convert.ToInt64(indexCmd.ExecuteScalar());
                if (exists == 0)
                {
                    Log("DB: idx_type index'i eksik, ekleniyor...");
                    var swIndex = Stopwatch.StartNew();
                    using (var addIndexCmd = conn.CreateCommand())
                    {
                        addIndexCmd.CommandTimeout = dbCommandTimeoutSeconds;
                        addIndexCmd.CommandText = "ALTER TABLE tbl_reports ADD INDEX idx_type (type)";
                        addIndexCmd.ExecuteNonQuery();
                    }
                    swIndex.Stop();
                    Log($"DB: idx_type index'i eklendi ({swIndex.ElapsedMilliseconds}ms).");
                }
            }

            using var tx = conn.BeginTransaction();

            var swDelete = Stopwatch.StartNew();
            using (var deleteCmd = conn.CreateCommand())
            {
                deleteCmd.Transaction = tx;
                deleteCmd.CommandTimeout = dbCommandTimeoutSeconds;
                deleteCmd.CommandText = "DELETE FROM tbl_reports WHERE type = @type";
                deleteCmd.Parameters.AddWithValue("@type", type);
                int deleted = deleteCmd.ExecuteNonQuery();
                Log($"DB: {deleted} eski kayıt silindi. Type: {type} ({swDelete.ElapsedMilliseconds}ms)");
            }
            swDelete.Stop();

            var swInsert = Stopwatch.StartNew();
            using (var insertCmd = conn.CreateCommand())
            {
                insertCmd.Transaction = tx;
                insertCmd.CommandTimeout = dbCommandTimeoutSeconds;
                insertCmd.CommandText = "INSERT INTO tbl_reports (branch_code, img, upload_date, type) VALUES (@branch_code, @img, @upload_date, @type)";
                var pBranch = insertCmd.Parameters.Add("@branch_code", MySqlDbType.VarChar);
                var pImg = insertCmd.Parameters.Add("@img", MySqlDbType.LongText);
                var pUploadDate = insertCmd.Parameters.Add("@upload_date", MySqlDbType.DateTime);
                var pType = insertCmd.Parameters.Add("@type", MySqlDbType.Int32);
                insertCmd.Prepare();

                int inserted = 0;
                foreach (var (branchCode, base64) in items)
                {
                    pBranch.Value = branchCode;
                    pImg.Value = base64;
                    pUploadDate.Value = DateTime.Now;
                    pType.Value = type;
                    insertCmd.ExecuteNonQuery();

                    inserted++;
                    if (inserted % 10 == 0 || inserted == items.Count)
                        Log($"DB: {inserted}/{items.Count} kayıt insert edildi (geçen süre: {swInsert.ElapsedMilliseconds}ms)");
                }
            }
            swInsert.Stop();

            var swCommit = Stopwatch.StartNew();
            tx.Commit();
            swCommit.Stop();

            Log($"DB: {items.Count} kayıt yazıldı ve commit edildi. Type: {type} ({swCommit.ElapsedMilliseconds}ms)");

            return (swConnect.ElapsedMilliseconds, swCreate.ElapsedMilliseconds, swDelete.ElapsedMilliseconds, swInsert.ElapsedMilliseconds, swCommit.ElapsedMilliseconds);
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
