#!/usr/bin/env python3
"""
PPT/PPTX 转 PNG 图片工具
依赖: LibreOffice, PyMuPDF
安装: brew install libreoffice && pip3 install PyMuPDF

使用方法:
    python3 convert.py --input /path/to/file.pptx --output ./slides --zip
"""

import subprocess
import os
import sys
import argparse
from pathlib import Path
from typing import Optional
import fitz  # PyMuPDF


def check_dependencies():
    """检查依赖是否安装"""
    missing = []
    
    # 检查soffice
    result = subprocess.run(["which", "soffice"], capture_output=True)
    if result.returncode != 0:
        missing.append("LibreOffice (brew install libreoffice)")
    
    # 检查PyMuPDF
    try:
        import fitz
    except ImportError:
        missing.append("PyMuPDF (pip3 install PyMuPDF)")
    
    if missing:
        print("缺少依赖，请安装:")
        for m in missing:
            print(f"  - {m}")
        return False
    return True


def pptx_to_images(
    pptx_path: str,
    output_dir: str = "slides",
    dpi_scale: float = 2.0
) -> bool:
    """
    将PPTX转换为PNG图片
    
    Args:
        pptx_path: PPTX文件路径
        output_dir: 输出目录
        dpi_scale: DPI缩放倍数 (2.0 ≈ 150dpi)
    
    Returns:
        bool: 是否成功
    """
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    temp_pdf = Path("/tmp") / f"{pptx_path.stem}.pdf"
    
    if not pptx_path.exists():
        print(f"错误: 文件不存在: {pptx_path}")
        return False
    
    if pptx_path.suffix.lower() not in (".pptx", ".ppt"):
        print(f"错误: 请输入PPT/PPTX文件")
        return False
    
    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"输入: {pptx_path}")
    print(f"输出: {output_dir}")
    
    # Step 1: PPTX -> PDF
    print("\n[1/2] 转换为PDF...")
    cmd_pdf = [
        "soffice", "--headless", "--norestore",
        "--convert-to", "pdf",
        "--outdir", "/tmp",
        str(pptx_path)
    ]
    
    result = subprocess.run(cmd_pdf, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"PDF转换失败: {result.stderr}")
        return False
    
    if not temp_pdf.exists():
        print(f"错误: PDF文件未生成")
        return False
    
    # Step 2: PDF -> PNG (每页)
    print("[2/2] 转换为PNG...")
    
    try:
        doc = fitz.open(temp_pdf)
        total_pages = len(doc)
        print(f"总页数: {total_pages}")
        
        for page_num in range(total_pages):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
            
            output_path = output_dir / f"{page_num + 1}.png"
            pix.save(str(output_path))
            print(f"  保存: {output_path.name}")
        
        doc.close()
        
    except Exception as e:
        print(f"PNG转换失败: {e}")
        return False
    finally:
        # 清理临时PDF
        if temp_pdf.exists():
            temp_pdf.unlink()
    
    # 统计
    files = list(output_dir.glob("*.png"))
    print(f"\n✅ 转换完成! 共生成 {len(files)} 张图片")
    
    return True


def create_zip(input_dir: str, output_name: str = "slides") -> Optional[str]:
    """将图片打包为ZIP"""
    import zipfile
    
    input_dir = Path(input_dir)
    zip_path = input_dir.parent / f"{output_name}.zip"
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for img in sorted(input_dir.glob("*.png")):
            zf.write(img, img.name)
    
    print(f"\n📦 已打包: {zip_path}")
    return str(zip_path)


def main():
    parser = argparse.ArgumentParser(description="PPT转PNG工具")
    parser.add_argument("--input", "-i", required=True, help="输入PPT文件路径")
    parser.add_argument("--output", "-o", default="slides", help="输出目录")
    parser.add_argument("--scale", "-s", type=float, default=2.0, help="DPI缩放 (默认2.0)")
    parser.add_argument("--zip", "-z", action="store_true", help="打包为ZIP")
    
    args = parser.parse_args()
    
    if not check_dependencies():
        sys.exit(1)
    
    success = pptx_to_images(args.input, args.output, args.scale)
    
    if success and args.zip:
        zip_path = create_zip(args.output, Path(args.input).stem)
        if zip_path:
            print(f"\n📎 可发送文件: {zip_path}")


if __name__ == "__main__":
    main()
