using Microsoft.AspNetCore.Mvc;
using Sebtum.Models;
using System.Data;

namespace Sebtum.Controllers
{
    public class ExcelController : Controller
    {
        private readonly ILogger<ExcelController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public ExcelController(IWebHostEnvironment IWebHostEnvironment)
        {
            _webHostEnvironment = IWebHostEnvironment;
        }


        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Vipisky()
        {
            // Инициализация настроек по умолчанию
            Настройки настройки = new Настройки();

            // Передача настроек в представление
            ViewData["Настройки"] = настройки;

            return View(настройки);
        }
        [HttpPost]
        public async Task<IActionResult> Vipisky(IFormFile file, Настройки настройки)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("Файл не был загружен.");
                }

                // Проверка типа файла
                if (!file.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest("Неверный тип файла. Допустим только Excel (.xlsx).");
                }

                string uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "Files");
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }
                string fileName = Path.GetFileName(file.FileName);
                string fileSavePath = Path.Combine(uploadsFolder, "Обработываемый_файл.xlsx");

                using (FileStream stream = new FileStream(fileSavePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }


                // Обработка файла
                Excel excel = new Excel();
                DataTable dataTable = excel.ReadExcelFile(@fileSavePath);



                Анализ_выписки анализ_Выписки = new Анализ_выписки();
                DataTable dtResult = анализ_Выписки.Сделать_анализ(dataTable, настройки);
                string processedFilePath = Path.Combine(_webHostEnvironment.WebRootPath, "processed");
                if (!Directory.Exists(processedFilePath))
                {
                    Directory.CreateDirectory(processedFilePath);
                }
                // Сохранение обработанных данных в новый файл
                processedFilePath = Path.Combine(processedFilePath, "Обработанный_файл.xlsx");
                excel.SaveExel(dtResult, processedFilePath);

                // Возвращение файла пользователю
                return File(System.IO.File.ReadAllBytes(processedFilePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Обработанный_файл.xlsx");



            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, "Ошибка при обработке файла Excel");
                return BadRequest("Произошла ошибка при обработке файла. Попробуйте снова.");
            }

        }
    }
}

