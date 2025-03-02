using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using System.Text;
using System.Threading.Tasks;
using Tesseract;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO; // Thêm namespace này để sử dụng Path
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

public class HomeController : Controller
{
    private readonly IWebHostEnvironment _environment;

    public HomeController(IWebHostEnvironment environment)
    {
        _environment = environment;
    }

    [HttpPost]
    public async Task<IActionResult> Upload(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            ViewBag.Message = "Vui lòng chọn file.";
            return View("Index");
        }

        string uploadsFolder = System.IO.Path.Combine(_environment.WebRootPath, "uploads");
        if (!Directory.Exists(uploadsFolder))
            Directory.CreateDirectory(uploadsFolder);

        string filePath = System.IO.Path.Combine(uploadsFolder, file.FileName);
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        return RedirectToAction("XuLi", new { fileName = file.FileName });
    }

    public string ExtractTextFromPdf(string filePath)
    {
        StringBuilder text = new StringBuilder();
        using (PdfReader reader = new PdfReader(filePath))
        {
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
            }
        }
        return text.ToString();
    }

    public string ExtractTextFromImage(string imagePath)
    {
        try
        {
            using (var engine = new TesseractEngine(@"C:\Program Files\Tesseract-OCR\tessdata", "eng", EngineMode.Default))
            using (var img = Pix.LoadFromFile(imagePath))
            using (var page = engine.Process(img))
            {
                string extractedText = page.GetText();
                if (string.IsNullOrEmpty(extractedText))
                {
                    Console.WriteLine("Không trích xuất được văn bản từ hình ảnh.");
                }
                return extractedText;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Lỗi khi trích xuất văn bản từ hình ảnh: {ex.Message}");
            return string.Empty;
        }
    }

    public async Task<string> CallAI(string text)
    {
        using (var client = new HttpClient())
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions");
            request.Headers.Add("Authorization", "AIzaSyBHqcZRBDR7F9_Y8wVh01gYbElLgeVaHyE"); // Thay YOUR_API_KEY bằng API key thực tế
// API key của em đâu? da em quen doi :((
            var jsonContent = new
            {
                model = "gpt-4",
                messages = new[]
                {
                new
                {
                    role = "user",
                    content = $"Hãy trích xuất thông tin từ đoạn văn sau và trả về danh sách JSON các món ăn, với format như sau:\n" +
                              "[{{\"Name\": \"Tên món\", \"Price\": \"Giá\"}}, ...]\n\n{text}"
                }
            },
                temperature = 0.5
            };

            request.Content = new StringContent(JsonConvert.SerializeObject(jsonContent), Encoding.UTF8, "application/json");

            var response = await client.SendAsync(request);
            string jsonResponse = await response.Content.ReadAsStringAsync();

            // Ghi phản hồi để kiểm tra dữ liệu từ OpenAI
            Console.WriteLine(jsonResponse);

            return jsonResponse; // Trả về JSON để xử lý tiếp
        }
    }


    public class MenuItem
    {
        public string Name { get; set; } = string.Empty;
        public string Price { get; set; } = string.Empty;
    }

    public void CreateExcelFile(List<MenuItem> menuItems, string filePath)
    {
        if (menuItems == null || menuItems.Count == 0)
        {
            throw new Exception("Danh sách món ăn rỗng, không thể tạo file Excel.");
        }

        using (ExcelPackage package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Menu");
            worksheet.Cells[1, 1].Value = "Tên món";
            worksheet.Cells[1, 2].Value = "Giá";

            int row = 2;
            foreach (var item in menuItems)
            {
                worksheet.Cells[row, 1].Value = item.Name;
                worksheet.Cells[row, 2].Value = item.Price;
                row++;
            }

            try
            {
                System.IO.File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
            catch (Exception ex)
            {
                throw new Exception("Lỗi khi ghi file Excel: " + ex.Message);
            }
        }
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult XuLi(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
        {
            ViewBag.Message = "Tên file không hợp lệ.";
            return View("Index");
        }
        ViewBag.FileName = fileName;
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> ChuyenDoi(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
        {
            return BadRequest("Tên file không hợp lệ.");
        }

        string uploadsFolder = System.IO.Path.Combine(_environment.WebRootPath, "uploads");
        string filePath = System.IO.Path.Combine(uploadsFolder, fileName);

        if (!System.IO.File.Exists(filePath))
        {
            return NotFound("File không tồn tại!");
        }

        try
        {
            // Trích xuất văn bản từ file (PDF hoặc hình ảnh)
            string extractedText = fileName.EndsWith(".pdf") ? ExtractTextFromPdf(filePath) : ExtractTextFromImage(filePath);

            if (string.IsNullOrEmpty(extractedText))
            {
                return BadRequest("Không thể trích xuất văn bản từ file.");
            }

            // Gọi OpenAI API để phân tích và nhận kết quả JSON
            string jsonResult = await CallAI(extractedText);

            // Ghi log phản hồi JSON để kiểm tra
            System.IO.File.WriteAllText(System.IO.Path.Combine(_environment.WebRootPath, "log.txt"), jsonResult); // Ghi JSON phản hồi vào file log trong thư mục webroot
            Console.WriteLine("Phản hồi từ OpenAI: " + jsonResult); // Ghi log ra console (xem trong terminal hoặc log file)

            List<MenuItem> menuItems;

            try
            {
                // Kiểm tra xem JSON có phải là đối tượng thay vì mảng không
                if (jsonResult.StartsWith("{"))
                {
                    var jobject = JsonConvert.DeserializeObject<JObject>(jsonResult);

                    // Kiểm tra xem phản hồi có đúng phần dữ liệu bạn cần không
                    var dataArray = jobject["choices"]?.FirstOrDefault()?["message"]?["content"];
                    if (dataArray == null)
                    {
                        return BadRequest("Không thể tìm thấy dữ liệu trong phản hồi.");
                    }

                    // Log dữ liệu phân tích để kiểm tra
                    Console.WriteLine("Dữ liệu phân tích: " + dataArray.ToString());
                    menuItems = JsonConvert.DeserializeObject<List<MenuItem>>(dataArray.ToString());
                }
                else
                {
                    menuItems = JsonConvert.DeserializeObject<List<MenuItem>>(jsonResult);
                }

                if (menuItems == null || menuItems.Count == 0)
                {
                    return BadRequest("Dữ liệu JSON không hợp lệ hoặc danh sách rỗng.");
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Lỗi khi phân tích dữ liệu từ OpenAI: " + ex.Message);
            }

            // Tạo thư mục tải xuống nếu chưa tồn tại
            string downloadsFolder = System.IO.Path.Combine(_environment.WebRootPath, "downloads");
            if (!Directory.Exists(downloadsFolder))
                Directory.CreateDirectory(downloadsFolder);

            // Đường dẫn tệp Excel
            string excelFilePath = System.IO.Path.Combine(downloadsFolder, $"{System.IO.Path.GetFileNameWithoutExtension(fileName)}.xlsx");

            // Tạo file Excel
            CreateExcelFile(menuItems, excelFilePath);

            if (!System.IO.File.Exists(excelFilePath))
            {
                return StatusCode(500, "Lỗi khi tạo file Excel!");
            }

            // Tải xuống tệp Excel ngay lập tức
            var memory = new MemoryStream();
            using (var stream = new FileStream(excelFilePath, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;

            // Trả về tệp Excel để người dùng tải xuống
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{System.IO.Path.GetFileNameWithoutExtension(fileName)}.xlsx");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Đã xảy ra lỗi: {ex.Message}");
        }
    }


}


