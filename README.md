# DocToPdfTool

文档转 PDF 工具，支持 WPS Office 和 Microsoft Office 双引擎，以及 HTML 文件转换。

## 功能

- **双引擎转换**：支持 WPS Office 和 Microsoft Office 两种转换引擎，运行时自由切换
- **格式支持**：Word（.doc/.docx）、Excel（.xls/.xlsx）、PPT（.ppt/.pptx）、文本（.txt/.rtf）、HTML（.html/.htm/.mhtml）
- **HTML 转换**：基于 Edge 浏览器 headless 模式，保留原始布局和样式，支持响应式页面
- **Excel 多工作表**：自动列宽适配、分页控制，完整转换所有工作表
- **单文件发布**：通过 Costura.Fody 将所有依赖打包为单个 exe 文件
- **单实例运行**：防止重复启动

## 使用

1. 选择源文件（支持多格式筛选）
2. 设置输出目录（默认与源文件同目录）
3. 选择转换引擎（WPS Office / Microsoft Office）
4. 点击"执行转换"
5. 转换成功后自动打开输出目录并选中生成的文件

## 构建

### Release 构建
```bash
dotnet build -c Release
```

### 单文件发布
```bash
publish_single.bat
```

输出在 `publish\SingleFile\DocToPdfTool.exe`

## 技术栈

- .NET Framework 4.8 / WPF
- COM 自动化（WPS / Office 后期绑定）
- Edge Chromium headless（HTML 转 PDF）
- Win32 API 窗口守卫（抑制 Office 弹窗）
- Costura.Fody（单文件打包）

## 致谢

本项目的 WPS Office 转换能力参考了 [WPSToPDF](https://github.com/lm3515/WPSToPDF)；Microsoft Office 转换部分参考了 [Pdfor](https://github.com/Vit-Lib/Pdfor)。感谢这些项目提供的灵感和参考。