# 📄 PPT to PNG — OpenClaw Skill

<p align="center">
    <picture>
        <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/openclaw-logo-text-dark.png">
        <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png">
        <img src="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png" alt="PPT to PNG" width="500">
    </picture>
</p>

<p align="center">
  <a href="https://github.com/QiaoTuCodes/ppt-to-png/actions"><img src="https://img.shields.io/github/actions/workflow/status/QiaoTuCodes/ppt-to-png?branch=main&style=for-the-badge" alt="CI status"></a>
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
# Clone to OpenClaw workspace
cp -r ppt-to-png ~/openclaw-workspace/skills/

# Install dependencies
brew install libreoffice
pip3 install PyMuPDF --break-system-packages
```

## 🚀 Quick Start

```bash
python3 convert.py --input /path/to/file.pptx --output ./output --zip
```

## 🔧 Requirements

- Python 3.8+
- LibreOffice (headless mode)
- PyMuPDF

## 📂 Project Structure

```
ppt-to-png/
├── SKILL.md
├── convert.py
├── README.md
├── README-CN.md
└── LICENSE
```

## 📜 License

MIT License

## 👥 Authors

- **魏然** - [GitHub](https://github.com/QiaoTuCodes)

---

<p align="center">
  <sub>Built with ❤️ for OpenClaw</sub>
</p>
