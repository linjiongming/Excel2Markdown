# Tests 文件夹

这个文件夹包含 Excel2Markdown 项目的测试工具和示例文件。

## 文件说明

### TestFileCreator 项目
- **Program.cs**: 用于创建测试用的 Excel 文件的工具
- **TestFileCreator.csproj**: 测试文件创建器的项目配置文件

## 使用方法

### 创建测试 Excel 文件
```bash
# 进入测试目录
cd Tests

# 运行测试文件创建器
dotnet run --project TestFileCreator.csproj
```

这将在项目根目录创建一个名为 `test.xlsx` 的测试文件，包含示例数据。

## 注意事项

- 测试文件会使用 EPPlus 库创建 Excel 文件
- 需要设置 EPPlus 的非商业许可证上下文
- 生成的测试文件可以用于验证主项目的转换功能

## 扩展

您可以在这个文件夹中添加更多的测试工具，例如：
- 批量测试脚本
- 性能测试工具
- 数据验证工具
- 自动化测试脚本