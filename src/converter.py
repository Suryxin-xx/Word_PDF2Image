"""
核心转换模块 — PDF / Word 导出为图片

支持格式: PNG / JPEG / TIFF / BMP / WEBP
支持扫描件增强（锐化 + 对比度 + 自动色阶）
"""

import os
import tempfile
from pathlib import Path
from typing import Callable, Optional

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

ProgressCB = Callable[[int, int, str], None]

# 格式 → (Pillow format, 文件后缀, 说明)
SUPPORTED_FORMATS = {
    "PNG":  ("PNG",  ".png",  "无损，适合截图/存档"),
    "JPEG": ("JPEG", ".jpg",  "有损压缩，适合照片（可调质量）"),
    "TIFF": ("TIFF", ".tif",  "无损，适合印刷/出版"),
    "BMP":  ("BMP",  ".bmp",  "无损，文件较大"),
    "WEBP": ("WEBP", ".webp", "Google 格式，平衡质量与大小"),
}

FORMAT_KEYS = list(SUPPORTED_FORMATS.keys())
DEFAULT_FORMAT = "PNG"
SUPPORTED_INPUT_EXTENSIONS = {".pdf", ".doc", ".docx"}


def parse_page_range(page_range: Optional[str], total: int) -> list[int]:
    """解析页码范围，返回 0-based 页码列表；None 表示全部页面。"""
    if total <= 0:
        raise ValueError("文档没有可导出的页面")
    if not page_range:
        return list(range(total))

    import re

    pages = []
    parts = re.split(r"[,，\s]+", page_range.strip())
    for part in parts:
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                start, end = int(a.strip()), int(b.strip())
                if start < 1 or end > total or start > end:
                    raise ValueError
                pages.extend(range(start - 1, end))
            else:
                page_num = int(part)
                if page_num < 1 or page_num > total:
                    raise ValueError
                pages.append(page_num - 1)
        except ValueError as exc:
            raise ValueError(
                f"页码范围无效: {part}（文档共 {total} 页）"
            ) from exc

    if not pages:
        raise ValueError("未解析到有效页码")

    # 去重并保持输入顺序
    return list(dict.fromkeys(pages))


def word_to_pdf(
    word_path: str,
    pdf_path: str,
    progress_cb: Optional[ProgressCB] = None,
) -> str:
    """
    通过本机 Microsoft Word 将 .doc/.docx 只读导出为 PDF。

    使用独立的隐藏 Word 实例，不保存或修改原文件；打开文档时强制禁用宏。
    """
    if os.name != "nt":
        raise RuntimeError("Word 转换目前仅支持 Windows")

    source = Path(word_path).resolve()
    target = Path(pdf_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Word 文件不存在: {source}")
    if source.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError(f"不支持的 Word 格式: {source.suffix}")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "缺少 Word 自动化组件，请先安装依赖: pip install pywin32"
        ) from exc

    if progress_cb:
        progress_cb(0, 0, "正在将 Word 转为 PDF...")

    target.parent.mkdir(parents=True, exist_ok=True)
    word = None
    document = None
    pythoncom.CoInitialize()
    try:
        try:
            # DispatchEx 创建独立实例，避免复用或关闭用户当前的 Word 窗口。
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            raise RuntimeError(
                "无法启动 Microsoft Word，请确认本机已安装桌面版 Word"
            ) from exc

        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass

        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            NoEncodingDialog=True,
        )
        document.ExportAsFixedFormat(
            OutputFileName=str(target),
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,   # wdExportOptimizeForPrint
        )
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Word 转 PDF 失败: {exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not target.is_file() or target.stat().st_size == 0:
        raise RuntimeError("Word 未能生成有效的临时 PDF")
    return str(target)


def document_to_images(
    input_path: str,
    output_dir: str,
    fmt: str = "PNG",
    dpi: int = 200,
    quality: int = 90,
    pages: Optional[list] = None,
    page_range: Optional[str] = None,
    progress_cb: Optional[ProgressCB] = None,
    image_enhance: bool = False,
    enhance_sharpness: int = 80,
    enhance_cutoff: int = 2,
    enhance_contrast: float = 1.15,
) -> list[str]:
    """
    将 PDF 或 Word 文档逐页导出为图片。

    Word 文件会先在系统临时目录中转为 PDF，完成或失败后自动清理。
    """
    source = Path(input_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"输入文件不存在: {source}")

    ext = source.suffix.lower()
    if ext not in SUPPORTED_INPUT_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_INPUT_EXTENSIONS))
        raise ValueError(f"不支持的输入格式: {ext or '无扩展名'}，可选: {supported}")
    if pages is not None and page_range is not None:
        raise ValueError("pages 和 page_range 不能同时指定")

    def convert_pdf(pdf_path: str) -> list[str]:
        selected_pages = pages
        if selected_pages is None and page_range is not None:
            import fitz

            with fitz.open(pdf_path) as doc:
                selected_pages = parse_page_range(page_range, len(doc))

        return pdf_to_images(
            pdf_path,
            output_dir,
            fmt=fmt,
            dpi=dpi,
            quality=quality,
            pages=selected_pages,
            progress_cb=progress_cb,
            image_enhance=image_enhance,
            enhance_sharpness=enhance_sharpness,
            enhance_cutoff=enhance_cutoff,
            enhance_contrast=enhance_contrast,
            output_base_name=source.stem,
        )

    if ext == ".pdf":
        return convert_pdf(str(source))

    with tempfile.TemporaryDirectory(prefix="pdf2image_word_") as temp_dir:
        temp_pdf = Path(temp_dir) / f"{source.stem}.pdf"
        word_to_pdf(str(source), str(temp_pdf), progress_cb=progress_cb)
        return convert_pdf(str(temp_pdf))


def pdf_to_images(
    pdf_path: str,
    output_dir: str,
    fmt: str = "PNG",
    dpi: int = 200,
    quality: int = 90,
    pages: Optional[list] = None,
    progress_cb: Optional[ProgressCB] = None,
    image_enhance: bool = False,
    enhance_sharpness: int = 80,
    enhance_cutoff: int = 2,
    enhance_contrast: float = 1.15,
    output_base_name: Optional[str] = None,
) -> list[str]:
    """
    将 PDF 每页导出为图片。

    参数:
        pdf_path:        输入 PDF 路径
        output_dir:      输出目录
        fmt:             图片格式 (PNG/JPEG/TIFF/BMP/WEBP)
        dpi:             渲染 DPI
        quality:         JPEG/WEBP 质量 (1-100)
        pages:           指定页码列表 (0-based)，None=全部
        progress_cb:     进度回调 (current, total, stage)
        image_enhance:   是否开启扫描件增强
        enhance_sharpness: 锐化强度 (0-200，0=不锐化，默认80)
        enhance_cutoff:   去黄力度 (0-10，0=不去黄，默认2)
        enhance_contrast: 对比度 (1.0-2.0，默认1.15)
        output_base_name: 输出文件名前缀，默认使用 PDF 文件名

    返回:
        生成的文件路径列表
    """
    import fitz

    fmt_key = fmt.upper()
    if fmt_key not in SUPPORTED_FORMATS:
        raise ValueError(f"不支持的格式: {fmt}，可选: {', '.join(FORMAT_KEYS)}")

    pil_fmt, ext, _ = SUPPORTED_FORMATS[fmt_key]

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    doc = fitz.open(pdf_path)
    total = len(doc)

    if pages is None:
        pages = list(range(total))
    else:
        pages = [p for p in pages if 0 <= p < total]

    if not pages:
        doc.close()
        raise ValueError("没有需要处理的页面")

    os.makedirs(output_dir, exist_ok=True)
    base_name = output_base_name or Path(pdf_path).stem

    generated = []
    save_kwargs = _get_save_kwargs(pil_fmt, quality)

    for idx, page_num in enumerate(pages):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=mat)

        # 通过 Pillow 处理
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # ---- 扫描件增强 ----
        if image_enhance:
            img = _enhance_image(img, sharpness=enhance_sharpness,
                                 cutoff=enhance_cutoff, contrast=enhance_contrast)
            if progress_cb:
                progress_cb(idx + 1, len(pages), "增强中")

        out_name = f"{base_name}_p{page_num + 1:04d}{ext}"
        out_path = os.path.join(output_dir, out_name)
        img.save(out_path, format=pil_fmt, **save_kwargs)

        generated.append(out_path)

        if progress_cb:
            progress_cb(idx + 1, len(pages), "")

    doc.close()
    return generated


def _enhance_image(img: Image.Image,
                   sharpness: int = 80,
                   cutoff: int = 2,
                   contrast: float = 1.15) -> Image.Image:
    """
    对扫描件图片进行增强处理。

    参数:
        sharpness: 锐化强度 0-200，0=跳过锐化（默认80）
        cutoff:    去黄力度 0-10，0=跳过自动色阶（默认2）
        contrast:  对比度 1.0-2.0（默认1.15）
    """
    # 1. 锐化（sharpness=0 时跳过）
    if sharpness > 0:
        img = img.filter(ImageFilter.UnsharpMask(
            radius=1.0, percent=sharpness, threshold=3))

    # 2. 自动色阶去黄（cutoff=0 时跳过）
    #    preserve_tone=True → 三个通道一起拉伸，不破坏色彩关系
    if cutoff > 0:
        img = ImageOps.autocontrast(img, cutoff=cutoff, preserve_tone=True)

    # 3. 对比度增强
    if contrast > 1.0:
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(contrast)

    return img


def _get_save_kwargs(pil_fmt: str, quality: int = 90) -> dict:
    """根据格式返回 Pillow save 的额外参数"""
    kwargs = {}
    if pil_fmt in ("JPEG",):
        kwargs["quality"] = quality
        kwargs["optimize"] = True
        kwargs["progressive"] = True
    elif pil_fmt == "WEBP":
        kwargs["quality"] = quality
        kwargs["method"] = 6  # 编码质量 0-6，6=最好
    elif pil_fmt == "TIFF":
        kwargs["compression"] = "lzw"
    # PNG/BMP 使用默认参数
    return kwargs
