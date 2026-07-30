"""
PDF / Word 导出为图片 - 主入口

将 PDF 或 Word 的每一页导出为图片文件，支持多种图片格式。

用法:
    python main.py                    # GUI 模式
    python main.py input.pdf          # CLI 模式（默认 PNG, 200 DPI）
    python main.py input.docx         # Word 转图片（需要安装 Microsoft Word）
    python main.py input.pdf -f JPEG --dpi 300 -q 95
    python main.py input.pdf -f TIFF -p 1-5,8,10
"""

import sys
import os
import argparse


if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))

from src.gui import PDF2ImageApp
from src.converter import document_to_images, FORMAT_KEYS, DEFAULT_FORMAT


def _fmt_size(byte: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if byte < 1024:
            return f"{byte:.1f} {unit}"
        byte /= 1024
    return f"{byte:.1f} TB"


def cli_mode():
    parser = argparse.ArgumentParser(
        description="将 PDF 或 Word 文档导出为图片"
    )
    parser.add_argument("input", help="PDF、DOC 或 DOCX 文件路径")
    parser.add_argument("-o", "--output", help="输出目录（默认: 文件名_图片/）")
    parser.add_argument("-f", "--format", default=DEFAULT_FORMAT,
                        choices=FORMAT_KEYS,
                        help=f"图片格式（默认: {DEFAULT_FORMAT}）")
    parser.add_argument("--dpi", type=int, default=200, help="渲染 DPI（默认: 200）")
    parser.add_argument("-q", "--quality", type=int, default=90,
                        help="JPEG/WEBP 质量 1-100（默认: 90）")
    parser.add_argument("-p", "--pages", help="页面范围，如: 1,3,5-10（默认: 全部）")
    parser.add_argument("--enhance", action="store_true",
                        help="开启扫描件增强（锐化 + 自动色阶 + 对比度）")
    parser.add_argument("--enhance-sharpness", type=int, default=80,
                        help="锐化强度 0-200（默认 80，0=不锐化）")
    parser.add_argument("--enhance-cutoff", type=int, default=2,
                        help="去黄力度 0-10（默认 2，0=不去黄）")
    parser.add_argument("--enhance-contrast", type=float, default=1.15,
                        help="对比度 1.0-2.0（默认 1.15）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"[错误] 文件不存在: {args.input}")
        sys.exit(1)

    # 输出目录
    from pathlib import Path
    inp = Path(args.input)
    out_dir = args.output or str(inp.parent / f"{inp.stem}_图片")
    os.makedirs(out_dir, exist_ok=True)

    def cb(current, total, stage):
        if total > 0:
            print(f"\r  进度: {current}/{total} 张 ({int(current/total*100)}%)", end="")
        sys.stdout.flush()

    try:
        print(f"文件:   {args.input}")
        print(f"输出:   {out_dir}")
        print(f"格式:   {args.format}")
        print(f"DPI:    {args.dpi}")
        if args.format in ("JPEG", "WEBP"):
            print(f"质量:   {args.quality}")
        if args.enhance:
            print(f"增强:   已开启")
            print(f"  锐化强度: {args.enhance_sharpness}")
            print(f"  去黄力度: {args.enhance_cutoff}")
            print(f"  对比度:   {args.enhance_contrast}")
        if args.pages:
            print(f"页面:   {args.pages}（自定义）")
        else:
            print("页面:   全部")

        generated = document_to_images(
            args.input, out_dir,
            fmt=args.format,
            dpi=args.dpi,
            quality=args.quality,
            page_range=args.pages,
            progress_cb=cb,
            image_enhance=args.enhance,
            enhance_sharpness=args.enhance_sharpness,
            enhance_cutoff=args.enhance_cutoff,
            enhance_contrast=args.enhance_contrast,
        )
        print()
        total_size = sum(os.path.getsize(f) for f in generated)
        print(f"\n[完成] 共导出 {len(generated)} 张图片，总大小 {_fmt_size(total_size)}")
        print(f"  目录: {out_dir}")
    except Exception as e:
        print(f"\n[错误] {e}")
        sys.exit(1)


def main():
    if len(sys.argv) > 1:
        cli_mode()
    else:
        app = PDF2ImageApp()
        app.run()


if __name__ == "__main__":
    main()
