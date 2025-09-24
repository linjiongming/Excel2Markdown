using OfficeOpenXml;
using System.Text;
using ExcelDataReader;
using System.Data;

namespace Excel2Markdown
{
    public class ExcelToMarkdownConverter
    {
        public static string ConvertToMarkdown(string excelFilePath)
        {
            var extension = Path.GetExtension(excelFilePath).ToLower();
            
            if (extension == ".xlsx")
            {
                return ConvertXlsxToMarkdown(excelFilePath);
            }
            else if (extension == ".xls")
            {
                return ConvertXlsToMarkdown(excelFilePath);
            }
            else
            {
                throw new NotSupportedException($"不支持的文件格式: {extension}");
            }
        }
        
        private static string ConvertXlsToMarkdown(string excelFilePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            var markdownBuilder = new StringBuilder();
            
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    
                    foreach (DataTable table in dataSet.Tables)
                    {
                        markdownBuilder.AppendLine($"## {table.TableName}");
                        markdownBuilder.AppendLine();
                        
                        if (table.Rows.Count == 0)
                        {
                            markdownBuilder.AppendLine("*工作表为空*");
                            markdownBuilder.AppendLine();
                            continue;
                        }
                        
                        // 找到实际有数据的列范围
                        int maxCol = 0;
                        int dataRowStart = 0;
                        bool hasData = false;
                        
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                var value = table.Rows[row][col];
                                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                                {
                                    if (!hasData)
                                    {
                                        dataRowStart = row;
                                        hasData = true;
                                    }
                                    if (col > maxCol) maxCol = col;
                                }
                            }
                        }
                        
                        if (!hasData)
                        {
                            markdownBuilder.AppendLine("*工作表为空*");
                            markdownBuilder.AppendLine();
                            continue;
                        }
                        
                        // 输出表头（假设第一行是表头）
                        var headers = new List<string>();
                        for (int col = 0; col <= maxCol; col++)
                        {
                            var headerValue = dataRowStart < table.Rows.Count ? 
                                FormatCellValue(table.Rows[dataRowStart][col]) : "";
                            headers.Add(headerValue);
                        }
                        
                        markdownBuilder.AppendLine("| " + string.Join(" | ", headers) + " |");
                        
                        // 输出分隔线
                        var separators = new string[headers.Count];
                        for (int i = 0; i < separators.Length; i++)
                        {
                            separators[i] = "---";
                        }
                        markdownBuilder.AppendLine("| " + string.Join(" | ", separators) + " |");
                        
                        // 输出数据行
                        for (int row = dataRowStart + 1; row < table.Rows.Count; row++)
                        {
                            var rowData = new List<string>();
                            for (int col = 0; col <= maxCol; col++)
                            {
                                var cellValue = FormatCellValue(table.Rows[row][col]);
                                cellValue = EscapeMarkdownChars(cellValue);
                                rowData.Add(cellValue);
                            }
                            
                            // 跳过完全空白的行
                            if (rowData.Any(cell => !string.IsNullOrWhiteSpace(cell)))
                            {
                                markdownBuilder.AppendLine("| " + string.Join(" | ", rowData) + " |");
                            }
                        }
                        
                        markdownBuilder.AppendLine();
                    }
                }
            }
            
            return markdownBuilder.ToString();
        }
        
        private static string ConvertXlsxToMarkdown(string excelFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var markdownBuilder = new StringBuilder();
            
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    markdownBuilder.AppendLine($"## {worksheet.Name}");
                    markdownBuilder.AppendLine();
                    
                    // 获取工作表的使用范围
                    if (worksheet.Cells.Any() == false)
                    {
                        markdownBuilder.AppendLine("*工作表为空*");
                        markdownBuilder.AppendLine();
                        continue;
                    }
                    
                    // 手动查找数据范围
                    int startRow = 1, startCol = 1;
                    int endRow = 1, endCol = 1;
                    
                    // 找到有数据的范围
                    for (int row = 1; row <= 1000; row++)
                    {
                        for (int col = 1; col <= 100; col++)
                        {
                            if (worksheet.Cells[row, col].Value != null)
                            {
                                if (row > endRow) endRow = row;
                                if (col > endCol) endCol = col;
                            }
                        }
                    }
                    
                    // 构建表头
                    var headers = new List<string>();
                    for (int col = startCol; col <= endCol; col++)
                    {
                        var headerValue = worksheet.Cells[startRow, col].Value?.ToString() ?? "";
                        headers.Add(headerValue);
                    }
                    
                    // 输出表头
                    markdownBuilder.AppendLine("| " + string.Join(" | ", headers) + " |");
                    
                    // 输出分隔线
                    var separators = new string[headers.Count];
                    for (int i = 0; i < separators.Length; i++)
                    {
                        separators[i] = "---";
                    }
                    markdownBuilder.AppendLine("| " + string.Join(" | ", separators) + " |");
                    
                    // 输出数据行
                    for (int row = startRow + 1; row <= endRow; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = startCol; col <= endCol; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                            // 转义 Markdown 特殊字符
                            cellValue = EscapeMarkdownChars(cellValue);
                            rowData.Add(cellValue);
                        }
                        markdownBuilder.AppendLine("| " + string.Join(" | ", rowData) + " |");
                    }
                    
                    markdownBuilder.AppendLine();
                }
            }
            
            return markdownBuilder.ToString();
        }
        
        private static string FormatCellValue(object value)
        {
            if (value == null)
                return "";
                
            // 处理DateTime类型
            if (value is DateTime dateTime)
            {
                // 如果日期在1900年附近，可能是Excel数字转换错误
                if (dateTime.Year <= 1900)
                {
                    // 尝试将其视为数字（天数）
                    var dayNumber = (dateTime - new DateTime(1900, 1, 1)).TotalDays + 1;
                    if (dayNumber > 0 && dayNumber <= 31) // 可能是日期
                    {
                        return dayNumber.ToString("0");
                    }
                    return dateTime.Day.ToString(); // 返回日
                }
                else if (dateTime.Year >= 2020 && dateTime.Year <= 2030)
                {
                    // 对于合理的日期，格式化为简单格式
                    return dateTime.ToString("M/d"); // 例如 1/1, 1/2
                }
            }
            
            // 处理数字类型
            if (value is double || value is float || value is decimal)
            {
                var numValue = Convert.ToDouble(value);
                // 如果是整数，不显示小数点
                if (numValue == Math.Floor(numValue))
                {
                    return numValue.ToString("0");
                }
                return numValue.ToString();
            }
            
            return value.ToString() ?? "";
        }
        
        private static string EscapeMarkdownChars(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
                
            // 转义 Markdown 表格中的管道符
            text = text.Replace("|", "\\|");
            
            // 转义其他可能导致问题的字符
            text = text.Replace("\n", "<br>");
            text = text.Replace("\r", "");
            
            return text;
        }
        
        public static void ConvertExcelFileToMarkdown(string excelFilePath)
        {
            try
            {
                Console.WriteLine($"正在处理文件: {Path.GetFileName(excelFilePath)}");
                
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine($"错误: 文件不存在 - {excelFilePath}");
                    return;
                }
                
                var markdownContent = ConvertToMarkdown(excelFilePath);
                
                // 生成输出文件路径（在相同目录下）
                var directory = Path.GetDirectoryName(excelFilePath) ?? "";
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelFilePath);
                var outputPath = Path.Combine(directory, $"{fileNameWithoutExtension}.md");
                
                // 写入 Markdown 文件
                File.WriteAllText(outputPath, markdownContent, Encoding.UTF8);
                
                Console.WriteLine($"成功生成: {Path.GetFileName(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文件 {Path.GetFileName(excelFilePath)} 时出错: {ex.Message}");
            }
        }
    }
}