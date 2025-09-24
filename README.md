# Excel2Markdown 使用说明

## 功能介绍
Excel2Markdown 是一个基于 .NET 的控制台应用程序，用于将 Excel 文件转换为 Markdown 格式的表格。

## 系统要求
- .NET 8.0 运行时环境（或更高版本）
- Windows 操作系统

## 支持的文件格式
- .xlsx (Excel 2007及以后版本)
- .xls (Excel 97-2003版本)

## 项目结构
```
Excel2Markdown/
├── Excel2Markdown/          # 主项目源代码
│   ├── Program.cs
│   ├── ExcelToMarkdownConverter.cs
│   └── Excel2Markdown.csproj
├── Tests/                   # 测试工具和示例
│   ├── Program.cs          # 测试文件创建器
│   ├── TestFileCreator.csproj
│   ├── RunTests.ps1        # 批量测试脚本
│   └── README.md
├── Publish/                 # 发布的可执行文件
│   └── Excel2Markdown.exe
└── README.md               # 项目说明
```

## 使用方法

### 方法1：拖拽操作（推荐）
1. 找到生成的 `Excel2Markdown.exe` 文件
2. 将一个或多个 Excel 文件直接拖拽到 `Excel2Markdown.exe` 图标上
3. 程序会自动处理所有拖拽的文件
4. 转换完成后，Markdown 文件将保存在与原 Excel 文件相同的目录中

### 方法2：命令行操作
打开命令提示符或 PowerShell，执行以下命令：

```bash
Excel2Markdown.exe "C:\path\to\your\file.xlsx"
```

也可以同时处理多个文件：

```bash
Excel2Markdown.exe "file1.xlsx" "file2.xlsx" "file3.xlsx"
```

## 输出格式
- 每个 Excel 工作表会转换为一个 Markdown 表格
- 工作表名称会作为二级标题显示
- 第一行会被视为表头
- 空工作表会显示 "*工作表为空*"
- 表格中的特殊字符会自动转义

## 示例

### 输入（Excel文件）
| 姓名 | 年龄 | 城市 |
|------|------|------|
| 张三 | 25   | 北京 |
| 李四 | 30   | 上海 |

### 输出（Markdown文件）
```markdown
## Sheet1

| 姓名 | 年龄 | 城市 |
| --- | --- | --- |
| 张三 | 25 | 北京 |
| 李四 | 30 | 上海 |
```

## 注意事项
1. 确保 Excel 文件没有被其他程序打开
2. 程序会自动覆盖同名的 Markdown 文件
3. 大文件可能需要较长处理时间
4. 建议在处理重要文件前先进行备份

## 错误处理
- 文件不存在：会显示错误信息并跳过该文件
- 文件被占用：会显示错误信息
- 不支持的文件格式：会自动跳过

## 开发和测试

### 编译项目
```bash
# 编译主项目
dotnet build .\Excel2Markdown\Excel2Markdown.csproj

# 发布单文件可执行程序
dotnet publish .\Excel2Markdown\Excel2Markdown.csproj --configuration Release --self-contained true -p:PublishSingleFile=true --runtime win-x64 --output .\Publish
```

### 运行测试
```bash
# 创建测试文件
dotnet run --project .\Tests\TestFileCreator.csproj

# 运行批量测试（PowerShell）
PowerShell -ExecutionPolicy Bypass -File ".\Tests\RunTests.ps1"
```

### 测试工具说明
- **TestFileCreator**: 创建包含多工作表的测试Excel文件
- **RunTests.ps1**: 自动化测试脚本，测试完整的转换流程
- 测试文件包含：员工信息表、项目统计表、空工作表（用于测试各种场景）

## 许可证
本程序使用 EPPlus 库（NonCommercial License）和 ExcelDataReader 库处理 Excel 文件。

---
如有问题，请检查：
1. 是否安装了 .NET 10.0 运行时
2. Excel 文件是否完整且未损坏
3. 是否有足够的磁盘空间保存输出文件