<div align="center">

# 🖼️ PDF / Word 导出为图片

**PDF & Word to Image Exporter** — 将 PDF、DOC、DOCX 文件的每一页导出为高清图片

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
[![Release](https://img.shields.io/github/v/release/Suryxin-xx/PDF2Image)](https://github.com/Suryxin-xx/PDF2Image/releases)

---

</div>

## 📌 简介

你是否遇到过这些场景？

- 需要从 PDF 中提取图片素材
- 想直接把 Word 文档逐页转换成图片
- 想把 PDF 的每一页拆成独立的图片文件
- 需要将 PDF 转为特定格式（如 JPEG / TIFF）用于打印或发布
- 扫描版 PDF 文字模糊，希望在导出时顺便增强清晰度

这个小工具就是为此而生，**将 PDF 或 Word 文档逐页导出为高清图片**，支持多种常见图片格式，内置扫描件增强功能。

## ✨ 功能特性

| 特性 | 说明 |
|------|------|
| 📄 **PDF / Word 输入** | 支持 PDF、DOC、DOCX；Word 文件通过本机 Microsoft Word 高保真渲染 |
| 🖼️ **5 种图片格式** | PNG / JPEG / TIFF / BMP / WEBP |
| 🎚️ **DPI 可调** | 150 / 200 / 300 / 400 DPI，平衡清晰度与文件大小 |
| ⚙️ **质量调节** | JPEG/WEBP 模式下可自定义压缩质量 |
| 📄 **页面范围选择** | 全部导出 / 自定义页码（如 `1,3,5-10`）|
| ✨ **扫描件增强** | 锐化 + 去尘 + 对比度调节，让老旧扫描件更清晰 |
| ⚡ **双模式运行** | **GUI 图形界面** + **CLI 命令行** |
| 📊 **实时进度** | 进度条 + 当前页码提示 |
| 📂 **一键直达** | 导出完成后直接打开输出文件夹 |

## 🖥️ 截图

![主界面截图](screenshots/ScreenShot.png)

*主界面：选择 PDF 或 Word → 增强设置 → 选择格式/DPI/页码 → 开始导出*

## 📦 下载

> 前往 [Releases](https://github.com/Suryxin-xx/PDF2Image/releases) 下载最新版 exe

| 文件 | 说明 |
|------|------|
| `PDF导出为图片.zip` | 单 exe 文件，解压即用（推荐） |

**系统要求：** Windows 10/11，64 位。转换 DOC/DOCX 时需要安装桌面版 Microsoft Word；仅转换 PDF 时不需要 Word。

## 🚀 使用方法

### GUI 模式（推荐）

双击运行 `PDF导出为图片.exe`：

1. **选择文档** — 点击“浏览”选择 PDF、DOC 或 DOCX 文件
2. **选择输出目录** — 图片保存到哪里
3. **启用增强（可选）** — 勾选"扫描件增强"并调节锐度/去尘/对比度参数
4. **选择格式** — PNG / JPEG / TIFF / BMP / WEBP
5. **调节 DPI 和质量** — JPEG/WEBP 时可调质量
6. **选择页面范围** — 全部或自定义
7. **点击导出** — 等待进度条走完

### 扫描件增强说明

针对老旧纸质版扫描 PDF，勾选 **扫描件增强** 后可调整：

| 参数 | 范围 | 默认值 | 作用 |
|------|------|--------|------|
| 锐化强度 | 0–200 | 80 | 增强文字边缘清晰度 |
| 去尘等级 | 0–10 | 2 | 去除扫描噪点和杂色 |
| 对比度 | 1.0–2.0 | 1.15 | 拉伸黑白对比，让文字更醒目 |

> 点击 **恢复默认** 可将三个参数一键复位。

### CLI 模式

```bash
# 基本用法（默认 PNG, 200 DPI）
PDF2Image.exe input.pdf

# Word 转图片（需要本机安装 Microsoft Word）
PDF2Image.exe input.docx

# 指定格式和 DPI
PDF2Image.exe input.pdf -f JPEG --dpi 300

# 指定质量（仅 JPEG/WEBP 有效）
PDF2Image.exe input.pdf -f JPEG -q 95

# 指定页面范围
PDF2Image.exe input.pdf -f TIFF -p 1-5,8,10

# 指定输出目录
PDF2Image.exe input.pdf -o D:\my_images

# 启用扫描件增强并自定义参数
PDF2Image.exe input.pdf --enhance --enhance-sharpness 120 --enhance-cutoff 3 --enhance-contrast 1.3
```

## 🔧 从源码运行

```bash
# 1. 克隆仓库
git clone https://github.com/Suryxin-xx/PDF2Image.git
cd PDF2Image

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行（GUI 模式）
python main.py

# 4. 运行（CLI 模式）
python main.py input.pdf -f JPEG --dpi 300

# Word 文档同样支持格式、DPI 和页码参数
python main.py input.docx -f PNG --dpi 300 -p 1-3
```

## 🏗️ 技术栈

| 组件 | 用途 |
|------|------|
| [Python](https://www.python.org/) | 编程语言 |
| [PyMuPDF (fitz)](https://pypi.org/project/PyMuPDF/) | PDF 渲染 |
| [Pillow](https://python-pillow.org/) | 图片编码、保存与增强处理 |
| [pywin32](https://pypi.org/project/pywin32/) | 调用本机 Microsoft Word，将 DOC/DOCX 后台导出为临时 PDF |
| [tkinter](https://docs.python.org/3/library/tkinter.html) | GUI 界面（内置） |
| [PyInstaller](https://pyinstaller.org/) | 打包为 exe |

Word 转换过程只读打开原文件，并在独立的隐藏 Word 实例中禁用宏。中间 PDF 位于系统临时目录，转换完成或失败后都会自动清理。

## 🗂️ 项目结构

```
PDF2Image/
├── src/                    # 源代码
│   ├── __init__.py
│   ├── converter.py        # 核心转换逻辑 + 图像增强
│   └── gui.py              # 图形界面
├── main.py                 # 入口（GUI / CLI 双模式）
├── requirements.txt        # Python 依赖
├── build_exe.ps1           # 打包脚本（需 PyInstaller + UPX）
├── screenshots/            # 截图
├── LICENSE                 # MIT 许可证
└── README.md
```

## 🔨 自行打包

```powershell
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行打包脚本
.\build_exe.ps1
```

打包后的 exe 位于 `dist/PDF2Image.exe`。

## 📄 许可证

本项目使用 [MIT License](LICENSE) — 欢迎 fork、修改、分发。

## 🤝 贡献

有问题或建议？欢迎提交 [Issue](https://github.com/Suryxin-xx/PDF2Image/issues) 或 Pull Request。

---

<div align="center">

**如果这个工具对你有帮助，欢迎 ⭐ Star 支持！**

</div>
