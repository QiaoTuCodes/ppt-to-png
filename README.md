<p align="center">
    <picture>
        <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png">
        <img src="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png" alt="PPT to PNG" width="500">
    </picture>
</p>

<p align="center">
  <strong>📄 PPT/PPTX to PNG Converter Skill for OpenClaw</strong>
</p>

<p align="center">
  <a href="https://github.com/QiaoTuCodes/ppt-to-png/releases"><img src="https://img.shields.io/github/v/release/QiaoTuCodes/ppt-to-png?include_prereleases&style=for-the-badge" alt="GitHub release"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge" alt="MIT License"></a>
  <a href="https://github.com/QiaoTuCodes/ppt-to-png/stargazers"><img src="https://img.shields.io/github/stars/QiaoTuCodes/ppt-to-png?style=for-the-badge" alt="Stars"></a>
</p>

将 PowerPoint 文件（.ppt/.pptx）转换为 PNG 图像格式，支持批量转换。

## ✨ Features

- 📄 **PPT/PPTX Conversion** - Convert PowerPoint to PNG images
- 📦 **Batch Processing** - Convert multiple slides at once
- 📨 **ZIP Packaging** - Auto-package to ZIP for easy sharing
- 🌐 **Feishu Ready** - Direct send via Feishu (飞书)

## 📦 Installation

```bash
# Clone this skill to your OpenClaw workspace
cp -r ppt-to-png ~/openclaw-workspace/skills/

# Install dependencies
brew install libreoffice
pip3 install PyMuPDF --break-system-packages
```

## 🚀 Quick Start

```python
import subprocess
import fitz  # PyMuPDF
import os
import zipfile

# Convert PPT to PDF
subprocess.run([
    "libreoffice", "--headless", "--convert-to", "pdf",
    "input.pptx", "--outdir", "output/"
])

# Convert PDF to PNG
doc = fitz.open("output.pdf")
for page_num in range(len(doc)):
    page = doc[page_num]
    pix = page.get_pixmap()
    pix.save(f"output/{page_num + 1}.png")
```

## CLI Usage

```bash
# Basic conversion
python3 convert.py --input /path/to/file.pptx --output ./output

# With ZIP packaging
python3 convert.py --input /path/to/file.pptx --output ./output --zip

# Specify output folder
python3 convert.py -i presentation.pptx -o ./images
```

## 📖 Documentation

- [English README](README.md)
- [中文文档](README-CN.md)
- [Skill Definition](SKILL.md)

## 🔧 Requirements

- Python 3.8+
- LibreOffice (headless mode)
- PyMuPDF (fitz)

## 📂 Project Structure

```
ppt-to-png/
├── SKILL.md           # OpenClaw skill definition
├── convert.py         # Main Python module
├── README.md          # English documentation
├── README-CN.md       # Chinese documentation
├── LICENSE            # MIT License
└── .gitignore
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📜 License

MIT License - see [LICENSE](LICENSE) for details.

## 👥 Authors

- **魏然 (Weiran)** - [GitHub](https://github.com/QiaoTuCodes)
- **焱焱 (Yanyan)** - yanyan@3c3d77679723a2fe95d3faf9d2c2e5a65559acbc97fef1ef37783514a80ae453

## 🙏 Acknowledgments

- [LibreOffice](https://www.libreoffice.org/) - Open source office suite
- [PyMuPDF](https://pymupdf.readthedocs.io/) - Python PDF processing
- [OpenClaw](https://github.com/openclaw/openclaw) team

---

<p align="center">
  <sub>Built with ❤️ for the OpenClaw community</sub>
</p>
