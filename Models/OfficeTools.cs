
using Microsoft.Office.Interop.Word;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Data;
using Xceed.Words.NET;

namespace Testing
{
    static public class OfficeTools
    {
        static public class Table
        {
            // Метод для чтения таблиц из байтовых данных
            public static List<(System.Data.DataTable table, string sheetName)> TableReadFileBytes(byte[] fileBytes, string extension)
            {
                if (fileBytes == null || fileBytes.Length == 0)
                {
                    throw new ArgumentException("Данные файла отсутствуют или пусты.");
                }

                if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    return Read_XLS(new MemoryStream(fileBytes));
                }
                else if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return Read_XLSX(new MemoryStream(fileBytes));
                }
                else
                {
                    throw new NotSupportedException("Формат файла не поддерживается.");
                }
            }
            private static List<(System.Data.DataTable table, string sheetName)> Read_XLS(Stream stream)
            {
                var allTables = new List<(System.Data.DataTable table, string sheetName)>();
                IWorkbook workbook = new HSSFWorkbook(stream);

                for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                {
                    ISheet sheet = workbook.GetSheetAt(sheetIndex);
                    string sheetName = sheet.SheetName;

                    IRow headerRow = sheet.GetRow(0);
                    if (headerRow == null) continue;

                    int cellCount = headerRow.LastCellNum;
                    System.Data.DataTable dataTable = new System.Data.DataTable(sheetName);

                    Dictionary<string, int> columnNameCounts = new Dictionary<string, int>();
                    int emptyHeaderCount = 0;
                    for (int i = 0; i < cellCount; i++)
                    {
                        ICell cell = headerRow.GetCell(i);
                        if (cell != null)
                        {
                            string columnName = cell.ToString();
                            if (columnNameCounts.ContainsKey(columnName))
                            {
                                columnNameCounts[columnName]++;
                                columnName += $"_{columnNameCounts[columnName]}";
                            }
                            else
                            {
                                columnNameCounts[columnName] = 1;
                            }
                            dataTable.Columns.Add(columnName);
                        }
                        else
                        {
                            emptyHeaderCount++;
                            dataTable.Columns.Add($"EmptyCol{emptyHeaderCount}");
                        }
                    }

                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row != null)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int j = 0; j < cellCount; j++)
                            {
                                ICell cell = row.GetCell(j);
                                if (cell != null)
                                {
                                    dataRow[j] = cell.ToString();
                                }
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                    }

                    allTables.Add((dataTable, sheetName));
                }

                return allTables;
            }
            private static List<(System.Data.DataTable table, string sheetName)> Read_XLSX(Stream stream)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                List<(System.Data.DataTable table, string sheetName)> allTables = new List<(System.Data.DataTable, string)>();

                using (var package = new ExcelPackage(stream))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet.Dimension == null) continue;

                        var dataTable = new System.Data.DataTable(worksheet.Name);
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            dataTable.Columns.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
                        }

                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var dataRow = dataTable.NewRow();
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        allTables.Add((dataTable, worksheet.Name));
                    }
                }

                return allTables;
            }

            // Метод для записи таблиц в байты возвращает байты
            public static byte[] TableSaveFileBytes(List<(System.Data.DataTable table, string sheetName)> sheets, string extension)
            {
                if (sheets == null || sheets.Count == 0)
                {
                    throw new ArgumentException("Список листов пуст.");
                }

                if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    return Save_XLS(sheets);
                }
                else if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return Save_XLSX(sheets);
                }
                else
                {
                    throw new NotSupportedException("Формат файла не поддерживается.");
                }
            }
            private static byte[] Save_XLS(List<(System.Data.DataTable table, string sheetName)> tables)
            {
                using (var stream = new MemoryStream())
                {
                    IWorkbook workbook = new HSSFWorkbook();

                    foreach (var (table, sheetName) in tables)
                    {
                        ISheet sheet = workbook.CreateSheet(sheetName);

                        IRow headerRow = sheet.CreateRow(0);
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            ICell cell = headerRow.CreateCell(i);
                            cell.SetCellValue(table.Columns[i].ColumnName);
                        }

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            IRow row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                ICell cell = row.CreateCell(j);
                                cell.SetCellValue(table.Rows[i][j].ToString());
                            }
                        }
                    }

                    workbook.Write(stream);
                    return stream.ToArray();
                }
            }
            private static byte[] Save_XLSX(List<(System.Data.DataTable table, string sheetName)> sheets)
            {
                using (var stream = new MemoryStream())
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (var package = new ExcelPackage())
                    {
                        foreach (var (dataTable, sheetName) in sheets)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(sheetName);

                            for (int col = 1; col <= dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                            }

                            for (int row = 0; row < dataTable.Rows.Count; row++)
                            {
                                for (int col = 0; col < dataTable.Columns.Count; col++)
                                {
                                    worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                                }
                            }
                        }

                        package.SaveAs(stream);
                    }

                    return stream.ToArray();
                }
            }


            //Метод для чтения таблиц в форматах  .xls  и .xlsx возвращает список листов
            public static List<(System.Data.DataTable table, string sheetName)> TableReadFilePath(string filePath)
            {

                if (filePath != null)
                {


                    if (filePath.EndsWith(".xls"))
                    {
                        return Read_XLS(filePath);
                    }
                    else if (filePath.EndsWith(".xlsx"))
                    {
                        return Read_XLSX(filePath);
                    }
                    else
                    {
                        throw new NotSupportedException("Формат файла не поддерживается.");
                    }
                }
                else
                {
                    throw new NotSupportedException("Путь к файлу отсутсвует.");
                }

            }
            private static List<(System.Data.DataTable table, string sheetName)> Read_XLS(string filePath)
            {
                try
                {
                    var allTables = new List<(System.Data.DataTable table, string sheetName)>();

                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook;

                        // Определяем тип файла по расширению
                        if (filePath.EndsWith(".xls"))
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }
                        else if (filePath.EndsWith(".xlsx"))
                        {
                            workbook = new XSSFWorkbook(fileStream);
                        }
                        else
                        {
                            throw new NotSupportedException("Формат файла не поддерживается");
                        }

                        // Перебираем все листы в книге
                        for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                        {
                            ISheet sheet = workbook.GetSheetAt(sheetIndex);
                            string sheetName = sheet.SheetName;

                            IRow headerRow = sheet.GetRow(0);

                            if (headerRow == null) // Обрабатываем пустые листы без ошибок
                            {
                                continue;
                            }

                            int cellCount = headerRow.LastCellNum;

                            // Создаем новый DataTable для текущего листа
                            System.Data.DataTable dataTable = new System.Data.DataTable(sheetName);

                            // Читаем заголовки столбцов (с возможными исправлениями)
                            Dictionary<string, int> columnNameCounts = new Dictionary<string, int>(); // Для отслеживания уникальных имен
                            int emptyHeaderCount = 0;
                            for (int i = 0; i < cellCount; i++)
                            {
                                ICell cell = headerRow.GetCell(i);
                                if (cell != null)
                                {
                                    string columnName = cell.ToString();

                                    // Проверяем на дублирующиеся имена и создаем уникальные имена при необходимости
                                    if (columnNameCounts.ContainsKey(columnName))
                                    {
                                        columnNameCounts[columnName]++;
                                        columnName += $"_{columnNameCounts[columnName]}";
                                    }
                                    else
                                    {
                                        columnNameCounts[columnName] = 1;
                                    }

                                    dataTable.Columns.Add(columnName);
                                }
                                else
                                {
                                    emptyHeaderCount++;
                                    dataTable.Columns.Add($"EmptyCol{emptyHeaderCount}");
                                }
                            }

                            // Читаем строки данных
                            for (int i = 1; i <= sheet.LastRowNum; i++)
                            {
                                IRow row = sheet.GetRow(i);
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int j = 0; j < cellCount; j++)
                                    {
                                        ICell cell = row.GetCell(j);
                                        if (cell != null)
                                        {
                                            dataRow[j] = cell.ToString();
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }

                            // Добавляем DataTable листа и имя листа в список allTables
                            allTables.Add((dataTable, sheetName));
                        }
                    }

                    return allTables;
                }
                catch (Exception)
                {

                    throw;
                }

            }
            private static List<(System.Data.DataTable table, string sheetName)> Read_XLSX(string filePath)
            {
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    List<(System.Data.DataTable table, string sheetName)> allTables = new List<(System.Data.DataTable, string)>();

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null) continue; // Пропускаем пустые листы

                            var dataTable = new System.Data.DataTable(worksheet.Name);

                            // Читаем заголовки
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                dataTable.Columns.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
                            }

                            // Читаем данные
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var dataRow = dataTable.NewRow();
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                                }
                                dataTable.Rows.Add(dataRow);
                            }

                            allTables.Add((dataTable, worksheet.Name));
                        }
                    }

                    return allTables;
                }
                catch (Exception)
                {

                    throw;
                }



            }


            //Метод для записи полученных листов в файл в формате .xls  и .xlsx
            public static void TableSaveFilePath(string filePath, List<(System.Data.DataTable, string)> sheets)
            {
                if (filePath != null && sheets != null)
                {

                    if (filePath.EndsWith(".xls"))
                    {
                        Save_XLSX(filePath, sheets);
                    }
                    else if (filePath.EndsWith(".xlsx"))
                    {
                        Save_XLSX(filePath, sheets);
                    }
                    else
                    {
                        throw new NotSupportedException("Формат файла не поддерживается");
                    }
                }
                else
                {
                    throw new NotSupportedException("Нет пути файла, либо лист с таблицами пуст");
                }


            }
            private static void Save_XLS(List<(System.Data.DataTable table, string sheetName)> tables, string filePath)
            {
                try
                {
                    IWorkbook workbook = new HSSFWorkbook(); // Создаем новый XLS файл

                    foreach (var (table, sheetName) in tables)
                    {
                        ISheet sheet = workbook.CreateSheet(sheetName);

                        // Создаем строку заголовков
                        IRow headerRow = sheet.CreateRow(0);
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            ICell cell = headerRow.CreateCell(i);
                            cell.SetCellValue(table.Columns[i].ColumnName);
                        }

                        // Заполняем строки данными
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            IRow row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                ICell cell = row.CreateCell(j);
                                cell.SetCellValue(table.Rows[i][j].ToString());
                            }
                        }
                    }

                    // Сохраняем файл
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fileStream);
                    }
                }
                catch (Exception)
                {

                    throw;
                }

            }
            private static void Save_XLSX(string filePath, List<(System.Data.DataTable table, string sheetName)> sheets)
            {
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (var package = new ExcelPackage())
                    {
                        foreach (var (dataTable, sheetName) in sheets)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(sheetName);

                            // Записываем заголовки
                            for (int col = 1; col <= dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                            }

                            // Записываем данные
                            for (int row = 0; row < dataTable.Rows.Count; row++)
                            {
                                for (int col = 0; col < dataTable.Columns.Count; col++)
                                {
                                    worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                                }
                            }
                        }

                        package.SaveAs(new FileInfo(filePath));
                    }
                }
                catch (Exception)
                {

                    throw;
                }


            }
        }

        static public class Word
        {
            //Чтение файла байтов
            public static string WordReadFileByte(byte[] fileBytes, string extension)
            {
                if (fileBytes != null && !string.IsNullOrEmpty(extension))
                {
                    switch (extension.ToLower())
                    {
                        case ".docx":
                            return ReadDocx(fileBytes);
                        case ".doc":
                            return ReadDoc(fileBytes);
                        default:
                            throw new NotSupportedException("Формат файла не поддерживается");
                    }
                }
                else
                {
                    throw new NotSupportedException("Некорректные входные данные");
                }
            }
            private static string ReadDocx(byte[] fileBytes)
            {
                using (MemoryStream memoryStream = new MemoryStream(fileBytes))
                {
                    using (var document = DocX.Load(memoryStream))
                    {
                        return document.Text;
                    }
                }
            }
            private static string ReadDoc(byte[] fileBytes)
            {
                object missing = System.Reflection.Missing.Value;
                object readOnly = true;
                object filePath = Path.GetTempFileName();

                try
                {
                    // Записываем байты во временный файл
                    File.WriteAllBytes((string)filePath, fileBytes);

                    Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                    Document doc = word.Documents.Open(ref filePath, ref missing, ref readOnly);
                    string text = doc.Content.Text;
                    doc.Close();
                    word.Quit();
                    return text;
                }
                finally
                {
                    // Удаляем временный файл после использования
                    if (File.Exists((string)filePath))
                    {
                        File.Delete((string)filePath);
                    }
                }
            }
            // Чтение файла пути
            public static string WordReadFilePath(string filePath)
            {
                if (filePath != null)
                {
                    string extension = Path.GetExtension(filePath).ToLower();
                    switch (extension)
                    {
                        case ".docx":
                            return ReadDocx(filePath);
                        case ".doc":
                            return ReadDocx(filePath);
                        default:
                            throw new NotSupportedException("Формат файла не поддерживается");
                    }
                }
                else
                {
                    throw new NotSupportedException("Нет пути файла");
                }
            }
            private static string ReadDocx(string filePath)
            {
                using (var document = DocX.Load(filePath))
                {
                    return document.Text;
                }
            }

            //Записывает полученный файл и возвращает байты
            public static byte[] FillTemplateByte(byte[] templateBytes, Dictionary<string, string> values, string filePath)
            {
                if (templateBytes != null)
                {
                    string extension = Path.GetExtension(filePath).ToLower();
                    switch (extension)
                    {
                        case ".docx":
                            return FillTemplateDocxByte(templateBytes, values);
                        case ".doc":
                            return FillTemplateDocxByte(templateBytes, values);
                        default:
                            throw new NotSupportedException("Формат файла не поддерживается");
                    }
                }
                else
                {
                    throw new NotSupportedException("Некорректные входные данные");
                }
            }
            public static byte[] FillTemplateDocxByte(byte[] templateBytes, Dictionary<string, string> values)
            {
                if (templateBytes != null)
                {
                    using (MemoryStream templateStream = new MemoryStream(templateBytes))
                    {
                        using (var document = DocX.Load(templateStream))
                        {
                            foreach (var item in values)
                            {
                                if (item.Key.StartsWith("Image:"))
                                {
                                    string[] imagePaths = item.Value.Split(';');
                                    var key = item.Key.Substring(6); // Получаем ключ без префикса "Image:"
                                    var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Image:{key}}}}}"));
                                    if (paragraph != null)
                                    {
                                        foreach (var imagePath in imagePaths)
                                        {
                                            try
                                            {
                                                var image = document.AddImage(imagePath);
                                                var picture = image.CreatePicture();
                                                paragraph.AppendPicture(picture);
                                                paragraph.AppendLine(); // Добавляем новую строку для следующего изображения
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Ошибка вставки изображения: {ex.Message}");
                                            }
                                        }
                                        paragraph.ReplaceText($"{{{{Image:{key}}}}}", string.Empty);
                                    }
                                }
                                else if (item.Key.StartsWith("Link:"))
                                {
                                    // Вставка гиперссылок
                                    var key = item.Key.Substring(5); // Получаем ключ без префикса "Link:"
                                    var hyperlinkText = item.Value; // Предполагаем, что значение item.Value содержит ссылку
                                    var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Link:{key}}}}}"));
                                    if (paragraph != null)
                                    {
                                        var hyperlink = document.AddHyperlink(hyperlinkText, new Uri(hyperlinkText));
                                        paragraph.ReplaceText($"{{{{Link:{key}}}}}", string.Empty);
                                        paragraph.AppendHyperlink(hyperlink);
                                    }
                                }
                                else if (item.Key.StartsWith("Table:"))
                                {
                                    // Вставка таблицы
                                    var tableData = item.Value.Split(';');
                                    var key = item.Key.Substring(6); // Получаем ключ без префикса "Table:"
                                    var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Table:{key}}}}}"));
                                    if (paragraph != null)
                                    {
                                        var table = document.AddTable(tableData.Length, tableData[0].Split(',').Length);
                                        for (int i = 0; i < tableData.Length; i++)
                                        {
                                            var rowData = tableData[i].Split(',');
                                            for (int j = 0; j < rowData.Length; j++)
                                            {
                                                table.Rows[i].Cells[j].Paragraphs[0].Append(rowData[j]);
                                            }
                                        }
                                        paragraph.InsertTableAfterSelf(table);
                                        paragraph.ReplaceText($"{{{{Table:{key}}}}}", string.Empty);
                                    }
                                }
                                else
                                {
                                    // Замена текста
                                    document.ReplaceText($"{{{{{item.Key}}}}}", item.Value);
                                }
                            }

                            using (MemoryStream outputStream = new MemoryStream())
                            {
                                document.SaveAs(outputStream);
                                return outputStream.ToArray();
                            }
                        }
                    }
                }
                else
                {
                    throw new NotSupportedException("Некорректные входные данные");
                }
            }

            //Записывает полученный файл по пути
            public static void FillTemplatePath(string filePath, Dictionary<string, string> values, string outputFilePath)
            {
                if (filePath != null)
                {
                    string extension = Path.GetExtension(filePath).ToLower();
                    switch (extension)
                    {
                        case ".docx":
                            FillTemplateDocxPath(filePath, values, outputFilePath);
                            break;
                        case ".doc":
                            FillTemplateDocxPath(filePath, values, outputFilePath);
                            break;
                        default:
                            throw new NotSupportedException("Формат файла не поддерживается");
                    }
                }
                else
                {
                    throw new NotSupportedException("Нет пути у файла");
                }
            }
            private static void FillTemplateDocxPath(string filePath, Dictionary<string, string> values, string outputFilePath)
            {
                const double maxWidth = 500;  // Установите максимальную ширину для изображения
                const double maxHeight = 500; // Установите максимальную высоту для изображения

                using (var document = DocX.Load(filePath))
                {
                    foreach (var item in values)
                    {
                        if (item.Key.StartsWith("Image:"))
                        {
                            string[] imagePaths = item.Value.Split(';');
                            var key = item.Key.Substring(6); // Получаем ключ без префикса "Image:"
                            var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Image:{key}}}}}"));
                            if (paragraph != null)
                            {
                                foreach (var imagePath in imagePaths)
                                {
                                    try
                                    {
                                        var image = document.AddImage(imagePath);
                                        var picture = image.CreatePicture();

                                        // Проверяем размеры изображения и масштабируем, если необходимо
                                        if (picture.Width > maxWidth || picture.Height > maxHeight)
                                        {
                                            double scalingFactor = Math.Min(maxWidth / picture.Width, maxHeight / picture.Height);
                                            picture.Width = (float)(picture.Width * scalingFactor);
                                            picture.Height = (float)(picture.Height * scalingFactor);

                                        }

                                        paragraph.AppendPicture(picture);
                                        paragraph.AppendLine(); // Добавляем новую строку для следующего изображения
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Ошибка вставки изображения: {ex.Message}");
                                    }
                                }
                                paragraph.ReplaceText($"{{{{Image:{key}}}}}", string.Empty);
                            }
                        }
                        else if (item.Key.StartsWith("Link:"))
                        {
                            // Вставка гиперссылок
                            var key = item.Key.Substring(5); // Получаем ключ без префикса "Link:"
                            var hyperlinkText = item.Value; // Предполагаем, что значение item.Value содержит ссылку
                            var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Link:{key}}}}}"));
                            if (paragraph != null)
                            {
                                var hyperlink = document.AddHyperlink(hyperlinkText, new Uri(hyperlinkText));
                                paragraph.ReplaceText($"{{{{Link:{key}}}}}", string.Empty);
                                paragraph.AppendHyperlink(hyperlink);
                            }
                        }
                        else if (item.Key.StartsWith("Table:"))
                        {
                            // Вставка таблицы
                            var tableData = item.Value.Split(';');
                            var key = item.Key.Substring(6); // Получаем ключ без префикса "Table:"
                            var paragraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains($"{{{{Table:{key}}}}}"));
                            if (paragraph != null)
                            {
                                var table = document.AddTable(tableData.Length, tableData[0].Split(',').Length);
                                for (int i = 0; i < tableData.Length; i++)
                                {
                                    var rowData = tableData[i].Split(',');
                                    for (int j = 0; j < rowData.Length; j++)
                                    {
                                        table.Rows[i].Cells[j].Paragraphs[0].Append(rowData[j]);
                                    }
                                }
                                paragraph.InsertTableAfterSelf(table);
                                paragraph.ReplaceText($"{{{{Table:{key}}}}}", string.Empty);
                            }
                        }
                        else
                        {
                            // Замена текста
                            document.ReplaceText($"{{{{{item.Key}}}}}", item.Value);
                        }
                    }
                    document.SaveAs(outputFilePath);
                }
            }

        }

        //Конвертирует таблицу в строку
        public static string ConvertDataTableToString(System.Data.DataTable table)
        {
            // Проверка на наличие данных в DataTable
            if (table == null || table.Rows.Count == 0)
            {
                throw new ArgumentException("DataTable должен содержать данные.");
            }

            // Создаем строку для хранения результата
            var result = new System.Text.StringBuilder();

            // Добавляем заголовки столбцов
            result.Append(string.Join(",", table.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            result.AppendLine();

            // Добавляем данные из таблицы
            foreach (DataRow row in table.Rows)
            {
                result.Append(string.Join(",", row.ItemArray));
                result.AppendLine();
            }

            return result.ToString().TrimEnd(); // Удаляем последний перевод строки
        }

        //Конфертирует список строк в одну большую строку
        public static string ListToString(List<string> dataList)
        {
            if (dataList == null)
            {
                throw new ArgumentNullException(nameof(dataList));
            }

            return string.Join(";", dataList);
        }
    }
}



