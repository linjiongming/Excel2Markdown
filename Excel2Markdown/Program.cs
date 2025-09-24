﻿﻿﻿﻿using Excel2Markdown;

if (args.Length == 0)
{
    Console.WriteLine("Excel2Markdown - Excel 文件转 Markdown 工具");
    Console.WriteLine("使用方法:");
    Console.WriteLine("  1. 将 Excel 文件拖拽到此程序上");
    Console.WriteLine("  2. 或者使用命令行: Excel2Markdown.exe <文件路径1> [文件路径2] ...");
    Console.WriteLine();
    Console.WriteLine("支持的文件格式: .xlsx, .xls");
    Console.WriteLine("生成的 Markdown 文件将保存在与 Excel 文件相同的目录中。");
    Console.WriteLine();
    Console.WriteLine("按任意键退出...");
    Console.ReadKey();
    return;
}

Console.WriteLine("Excel2Markdown 转换工具");
Console.WriteLine("========================");
Console.WriteLine();

int successCount = 0;
int totalCount = args.Length;

foreach (string filePath in args)
{
    if (string.IsNullOrWhiteSpace(filePath))
        continue;
        
    // 检查文件扩展名
    var extension = Path.GetExtension(filePath).ToLower();
    if (extension != ".xlsx" && extension != ".xls")
    {
        Console.WriteLine($"跳过不支持的文件: {Path.GetFileName(filePath)}");
        continue;
    }
    
    try
    {
        ExcelToMarkdownConverter.ConvertExcelFileToMarkdown(filePath);
        successCount++;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"处理文件失败 {Path.GetFileName(filePath)}: {ex.Message}");
    }
}

Console.WriteLine();
Console.WriteLine($"转换完成! 成功: {successCount}/{totalCount}");

if (successCount < totalCount)
{
    Console.WriteLine("按任意键退出...");
    Console.ReadKey();
}
