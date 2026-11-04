# DocToPdfTool

多功能文档格式转换工具，支持 **文档转 PDF**、**Word 转 PDF**、**Excel 转 PDF**、**Excel 转图片**、**PPT 转 PDF**、**PPT 转 Word**、**PPT 转图片**、**Word 转图片**、**图片转 PDF**、**PDF 转 Word**、**PDF 转 Excel**、**PDF 转 PPT**、**PDF 转图片**、**PDF 转扫描件**、**拆分 PDF** 和 **合并 PDF** 等多种转换功能。

![DocToPdfTool](icons/DocToPdfTool.png)

## 界面

- **首页**：程序启动后展示功能主页，顶部为搜索框（输入关键词即可模糊匹配功能），下方为所有功能按钮网格（水平居中，功能增多时可滚动）
- **功能页**：点击首页功能按钮或搜索结果后，在同一窗口内切换到「左侧导航 + 右侧功能区」布局，并自动定位左侧菜单选中项（自动滚动到可视区）
- **返回首页**：功能页左上角提供「🏠 首页」按钮，可随时返回主页
- **单窗口路由**：首页与功能页之间为同一窗口内的页面切换，不会打开新窗口

## 功能

### 文档转 PDF
- **双引擎转换**：支持 WPS Office 和 Microsoft Office 两种转换引擎，运行时自由切换
- **格式支持**：Word（.doc/.docx）、Excel（.xls/.xlsx）、PPT（.ppt/.pptx）、文本（.txt/.rtf）、HTML（.html/.htm/.mhtml）
- **HTML 转换**：基于 Edge 浏览器 headless 模式，保留原始布局和样式
- **Excel 多工作表**：不做额外页面设置改动，输出与 WPS / Office 自身「导出为 PDF」的默认打印效果一致，完整转换所有工作表
- **批量转换**：支持添加/拖拽多个文档，逐文件显示转换进度
- **拖拽添加**：支持拖拽文档到列表，列表中间显示操作提示

### Word 转 PDF
- **内置引擎，无需外部依赖**：基于 Spire.Doc 本地 DLL，无需安装 Microsoft Office 或 WPS
- **批量转换**：支持同时添加多个 Word 文件，逐个转换
- **拖拽添加**：支持拖拽 Word 文件到列表

### Excel 转 PDF
- **内置引擎，无需外部依赖**：基于 Spire.XLS 本地 DLL，无需安装 Microsoft Office 或 WPS
- **格式支持**：Excel（.xls/.xlsx）
- **三档页面缩放**：
  - `按宽度适应`（默认）— 所有列压缩到一页宽，行继续纵向分页，**列不会被横向切断**
  - `整表缩放到单页` — 每张工作表输出为一页，小表一页看完
  - `实际大小` — 完全按原表列宽分页，超宽时横向分页（还原 Excel 原生打印效果）
- **输出整洁**：转换结果不含附加标注页与提示文字（引擎自动生成的提示工作表会自动剔除）
- **内存优化**：转换前跳过公式重算、转换后自动清理引擎缓存，降低批量转换时的内存峰值
- **批量转换**：支持同时添加多个 Excel 文件，逐个转换
- **拖拽添加**：支持拖拽 Excel 文件到列表

### PPT 转 PDF
- **内置引擎，无需外部依赖**：基于 Spire.Presentation 本地 DLL，无需安装 Microsoft Office 或 WPS
- **格式支持**：PowerPoint（.ppt/.pptx）
- **长文档完整转换**：本地 DLL 引擎支持大页数文档完整输出，内容不截断
- **内存优化**：每次转换后强制 GC + 归还内存给操作系统，降低批量转换时的内存峰值
- **批量转换**：支持同时添加多个 PPT 文件，逐个转换
- **拖拽添加**：支持拖拽 PPT 文件到列表

### PPT 转 Word
- **内置引擎，无需外部依赖**：基于本地 DLL（Spire.Presentation + Spire.PDF），无需安装 Microsoft Office 或 WPS
- **格式支持**：PowerPoint（.ppt/.pptx）→ Word（.docx）
- **两段式转换**：先由 Spire.Presentation 将 PPT 渲染为中间 PDF，再由 Spire.PDF 转换为 DOCX，临时文件自动清理
- **完整输出**：中间 PDF 由本地 DLL 生成，最终 DOCX 完整保留全部页面
- **内存优化**：每次转换后强制 GC + 归还内存给操作系统
- **批量转换**：支持同时添加多个 PPT 文件，逐个转换
- **拖拽添加**：支持拖拽 PPT 文件到列表

### Excel 转图片
- **内置引擎，无需外部依赖**：本地 DLL 渲染，无需安装 Microsoft Office 或 WPS
- **每表一图**：每张有内容的工作表输出为一张完整图片，命名 `工作簿名_工作表名.png`
- **内容不截断**：整表按内容缩放到单页再渲染，一张表的内容完整呈现在一张图内，不会被分页切开
- **自动裁边**：图片尺寸贴合表格实际内容，不含 A4 纸张的多余空白边距
- **输出整洁**：转换结果不含附加标注文字与提示页
- **尺寸保护**：超大表（如数千行）超出位图上限时自动下调 DPI（下限 24），保证单图完整输出不崩溃
- **可调参数**：支持 DPI（96/150/300，默认 150）、输出格式（JPG/PNG，默认 PNG 无损）
- **临时文件自动清理**：逐表中转 PDF 与分页图片在转换结束后删除（带句柄重试）
- **批量转换**：支持同时添加多个 Excel 文件，逐个转换
- **拖拽添加**：支持拖拽 Excel 文件到列表

### PPT 转图片
- **内置引擎，无需外部依赖**：基于 Spire.Presentation 本地 DLL，无需安装 Microsoft Office 或 WPS
- **逐页导出**：每张幻灯片输出为独立图片，命名 `原文件名_1.png`、`原文件名_2.png` …
- **长文档完整转换**：本地 DLL 引擎支持大页数文档完整输出，内容不截断
- **高质量渲染**：开启 `HighQualityImage` 高质量渲染，JPG 使用 100% 质量编码
- **可调参数**：支持 DPI（96/150/300）、输出格式（JPG/PNG，默认 PNG 无损）
- **内存优化**：每页完成即触发回调回收，转换结束强制 GC + 归还内存
- **批量转换**：支持同时添加多个 PPT 文件，逐个转换
- **拖拽添加**：支持拖拽 PPT 文件到列表

### Word 转图片
- **内置引擎，无需外部依赖**：基于 Spire.Doc 本地 DLL，无需安装 Microsoft Office 或 WPS
- **矢量重渲染**：通过 `SaveToImages(ImageType.Metafile)` 取得矢量 EMF，再按目标 DPI 绘制，文字/表格完整且任意 DPI 下均清晰
- **长文档完整转换**：本地 DLL 引擎支持大页数文档完整输出，内容不截断
- **可调参数**：支持 DPI（96/150/300）、输出格式（JPG/PNG，默认 PNG 无损）
- **无 Office 副作用**：不启动 Word 进程，无窗口弹出，单文件耗时远低于 COM 方案
- **批量转换**：支持同时添加多个 Word 文件，逐个转换
- **拖拽添加**：支持拖拽 Word 文件到列表

### PDF 转 Word
- **双引擎转换**：支持 Microsoft Office 和内置引擎（Spire.PDF）两种转换引擎，运行时自由切换
- **内置引擎无需 Office**：基于本地 PDF 引擎，无需安装 Microsoft Office，长文档完整转换
- **大文件分批处理**：超过 20 页的 PDF 自动分批转换后合并，内存峰值可控
- **Microsoft Office 引擎**：基于 Word COM 自动化，质量高但需要安装 Office
- **批量转换**：支持同时添加多个 PDF 文件，逐个转换
- **拖拽添加**：支持拖拽 PDF 文件到列表

### PDF 转 Excel
- **PdfPig 解析引擎**：基于 PdfPig 开源 PDF 文本提取库，不依赖 Office 组件
- **表格自动识别**：智能检测页面的表格结构，自动识别列边界和表头
- **文本换行合并**：自动合并因单元格内文本换行而被拆分的行，还原表格原始结构
- **页眉页脚过滤**：自动去除 PDF 页面的页眉页脚干扰内容
- **非表格降级**：非表格页面自动降级为文本行输出，避免数据丢失
- **多页支持**：多页 PDF 每页生成独立 Sheet，单页则输出为 Sheet1
- **Excel 原生写入**：纯 XML 方式生成 `.xlsx` 文件，无需 Excel 或任何 Office 组件
- **拖拽添加**：支持拖拽 PDF 文件到列表

### PDF 转 PPT
- **Windows 内置引擎**：基于 Windows 10/11 内置 PDF 渲染引擎（`Windows.Data.Pdf`），无需额外安装软件
- **图片转 PPT**：将 PDF 每页渲染为图片后嵌入 PPTX 幻灯片，保持原始排版
- **批量转换**：支持同时添加多个 PDF 文件，每个文件生成独立 PPTX
- **拖拽添加**：支持拖拽 PDF 文件到列表
- **零依赖**：不依赖 pdfium.dll、Ghostscript 或 Office 组件
- **兼容性好**：生成的 PPTX 在 WPS 和 Microsoft Office PowerPoint 中均可正常打开

### PDF 转图片
- **Windows 内置引擎**：基于 Windows 10/11 内置 PDF 渲染引擎（`Windows.Data.Pdf`），无需额外安装软件
- **批量转换**：支持同时添加多个 PDF 文件，每页生成一张图片
- **拖拽添加**：支持拖拽 PDF 文件到列表
- **可调参数**：支持 DPI（72/96/150，默认 150）、缩放倍率（1×~5×，默认 1×）、输出格式（JPG/PNG）
- **像素上限保护**：单页超过 2.5 亿像素时自动等比缩小，避免 GDI+ 位图内存溢出
- **零依赖**：不依赖 pdfium.dll、Ghostscript 或 Office 组件

### 图片转 PDF
- **零依赖**：纯 .NET 直接写入 PDF 1.4（`BinaryWriter` + DCTDecode），不依赖任何 PDF 库或 Office 组件
- **两种输出方式**：
  - `合成一个 PDF`（默认）— 所有图片按列表顺序合成一个 PDF，一张图一页
  - `每张图各出一个 PDF` — 每张图片单独生成一个 PDF，命名 `文件名_001.pdf`、`文件名_002.pdf` …
- **JPEG 无损嵌入**：普通 JPEG 原图以 DCTDecode 直接嵌入，不重编码，画质零损失、速度更快
- **格式转换**：PNG / BMP / GIF / TIFF 自动铺白底后转 JPEG 嵌入（PDF 不支持透明通道）
- **EXIF 方向纠正**：自动读取手机竖拍照片的方向标记并旋转到正确朝向
- **三档页面尺寸**：
  - `A4 适应`（默认）— 等比缩放放进 A4，按图片横竖自动选择纸张方向，四周留白
  - `A4 铺满` — 等比放大填满 A4，超出部分居中裁切
  - `原始尺寸` — 按图片自身 DPI 换算物理尺寸
- **边距可选**：0 / 5 / 10 / 20 mm（默认 10 mm）
- **顺序调整**：列表顺序即页序，支持上移 / 下移 / 移除，也支持拖拽一次追加多张
- **缩略图预览**：列表显示缩略图与原始像素尺寸，便于核对顺序
- **大图保护**：长边超过 5000 px 自动降采样，避免解码内存峰值过高
- **容错**：单张图片损坏或格式不支持时跳过并提示，不中断整个任务
- **重名保护**：输出文件重名时自动追加序号

### PDF 转扫描件
- **Windows 内置引擎**：基于 Windows 10/11 内置 PDF 渲染引擎（`Windows.Data.Pdf`），无需额外安装软件
- **图片型 PDF**：将 PDF 每页渲染为图片后嵌入新 PDF，文字不可选取编辑
- **批量转换**：支持同时添加多个 PDF 文件，每个文件生成独立扫描件 PDF
- **拖拽添加**：支持拖拽 PDF 文件到列表
- **零依赖**：纯 .NET 写入 PDF，无需 pdfium.dll、Ghostscript 或 Office 组件

### 拆分 PDF
- **按页数拆分**：每 N 页拆分为一个文件（默认每 1 页）
- **按范围拆分**：自定义页码范围，支持单页与范围混合，如 `1-5,6,7-8`
- **页数预览**：添加文件后自动显示每个 PDF 的页数，方便填写拆分参数
- **零依赖**：基于 PdfPig 的 PdfDocumentBuilder 逐页复制，保留原始文字、矢量和图片，不会栅格化
- **批量拆分**：支持同时添加多个 PDF 文件

### 合并 PDF
- **顺序合并**：列表顺序即合并顺序，支持「上移 / 下移」调整文件次序
- **页数预览**：添加文件后自动显示每个 PDF 的页数
- **防覆盖**：输出文件与已有文件重名时自动追加序号
- **零依赖**：基于 PdfPig 的 PdfDocumentBuilder 逐页复制，保留原始质量，不会栅格化

## 设置

- 点击右上角 **「≡ 设置」** 按钮打开设置窗口
- 可查看和修改**默认保存目录**（未配置时默认取系统「下载」目录）
- 设置保存在 `%APPDATA%\DocToPdfTool\config.ini`，永久生效

所有转换结果默认输出到「下载目录」，并按功能名自动创建子目录存放（目录存在则不重复创建）：

| 功能 | 输出子目录 |
|------|-----------|
| 文档转 PDF | `下载目录\文档转PDF\` |
| Word 转 PDF | `下载目录\Word转PDF\` |
| Excel 转 PDF | `下载目录\Excel转PDF\` |
| Excel 转图片 | `下载目录\Excel转图片\` |
| PPT 转 PDF | `下载目录\PPT转PDF\` |
| PPT 转 Word | `下载目录\PPT转Word\` |
| PPT 转图片 | `下载目录\PPT转图片\` |
| Word 转图片 | `下载目录\Word转图片\` |
| PDF 转 Word | `下载目录\PDF转Word\` |
| PDF 转 Excel | `下载目录\PDF转Excel\` |
| PDF 转 PPT | `下载目录\PDF转PPT\` |
| PDF 转图片 | `下载目录\PDF转图片\` |
| 图片转 PDF | `下载目录\图片转PDF\` |
| PDF 转扫描件 | `下载目录\PDF转扫描件\` |
| 拆分 PDF | `下载目录\拆分PDF\` |
| 合并 PDF | `下载目录\合并PDF\` |

## 使用

1. 程序启动后进入**首页**，可点击功能按钮，或在顶部搜索框输入关键词（如 `PDF`、`PPT`）快速筛选功能
2. 进入功能页后，也可通过**左侧功能菜单**切换其他功能；点击左上角「🏠 首页」可返回主页
3. 点击「添加文件」按钮或直接拖拽文件到列表
4. 按需求选择转换引擎 / 拆分方式 / 图片参数
5. 点击「执行转换」
6. 转换完成后自动打开输出目录

> 输出目录可在设置窗口中修改默认保存目录后生效，各功能的子目录会自动创建。

## 构建

### Release 构建
```bash
dotnet build -c Release
```

### 单目录发布
```bash
publish_dir.bat
```

输出在 `publish\DirPublish\`（exe + 所有依赖 DLL）

### 单文件发布
```bash
publish_single.bat
```

输出在 `publish\SingleFile\DocToPdfTool.exe`

## 技术栈

- .NET Framework 4.8 / WPF
- COM 自动化（WPS / Office 后期绑定）
- Edge Chromium headless（HTML 转 PDF）
- **Spire.PDF 12.10.7**（PDF 转 Word / PPT 转 Word 中段，本地 DLL 引用，无需 Office）
- **Spire.Doc 14.10.9**（Word 转 PDF / Word 转图片，本地 DLL 引用，无需 Office）
- **Spire.XLS 16.10.3**（Excel 转 PDF，以及 Excel 转图片的中间 PDF 生成；本地 DLL 引用，无需 Office）
- **Spire.Presentation 11.10.4**（PPT 转 PDF / PPT 转 Word / PPT 转图片，本地 DLL 引用，无需 Office）
- **矢量 EMF 重渲染**（Word 转图片：`Spire.Doc.SaveToImages(Metafile)` + GDI+ 按目标 DPI 绘制）
- **DocumentFormat.OpenXml**（DOCX 合并 / 内容清理）
- Word COM 自动化（PDF 转 Word，3 层回退创建真实 Word 实例）
- PdfPig 开源 PDF 文本提取（PDF 转 Excel、拆分 PDF（PdfDocumentBuilder）、合并 PDF）
- 纯 XML xlsx 生成（`ZipArchive` + `XElement`，零 Office 依赖）
- Win32 API 窗口守卫（抑制 Office 弹窗）
- WMI 进程监控 + WinEventHook（Word 窗口自动隐藏）
- Windows.Data.Pdf 内置 PDF 渲染引擎（PDF 转图片/PPT/扫描件、Excel 转图片，Windows 10+）
- 纯 XML pptx 生成（`ZipArchive` + `XElement`，图片嵌入幻灯片，零 Office 依赖）
- 纯 .NET PDF 写入（`BinaryWriter` + PDF 规范，零依赖嵌入 JPEG 图片；PDF 转扫描件 / 图片转 PDF 共用）
- ini 配置（`%APPDATA%\DocToPdfTool\config.ini`，默认保存目录持久化）
- Costura.Fody（单文件打包）

## 致谢

本项目的 WPS Office 转换能力参考了 [WPSToPDF](https://github.com/lm3515/WPSToPDF)；Microsoft Office 转换部分参考了 [Pdfor](https://github.com/Vit-Lib/Pdfor)。感谢这些项目提供的灵感和参考。