<p align="center">
    <picture>
        <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png">
        <img src="https://raw.githubusercontent.com/QiaoTuCodes/ppt-to-png/main/assets/ppt-to-png-logo.png" alt="PPT转PNG" width="500">
    </picture>
</p>

<p align="center">
  <strong>📄 PPT/PPTX 转 PNG 技能 for OpenClaw</strong>
</p>

<p align="center">
  <a href="https://github.com/QiaoTuCodes/ppt-to-png/releases"><img src="https://img.shields.io/github/v/release/QiaoTuCodes/ppt-to-png?include_prereleases&style=for-the-badge" alt="GitHub release"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge" alt="MIT License"></a>
  <a href="https://github.com/QiaoTuCodes/ppt-to-png/stargazers"><img src="https://img.shields.io/github/stars/QiaoTuCodes/ppt-to-png?style=for-the-badge" alt="Stars"></a>
</p>

将 PowerPoint 文件（.ppt/.pptx）转换为 PNG 图像格式，支持批量转换。

## ✨ 功能特性

- 📄 **PPT/PPTX 转换** - 将 PowerPoint 转换为 PNG 图片
- 📦 **批量处理** - 一次性转换多页幻灯片
- 📨 **ZIP打包** - 自动打包为 ZIP 便于分享
- 🌐 **飞书集成** - 支持直接通过飞书发送

## 📦 安装

```bash
# 复制技能到 OpenClaw 工作区
cp -r ppt-to-png ~/openclaw-workspace/skills/

# 安装依赖
brew install libreoffice
pip3 install PyMuPDF --break-system-packages
```

## 🚀 快速开始

```python
import subprocess
import fitz  # PyMuPDF
import os
import zipfile

# PPT 转 PDF
subprocess.run([
    "libreoffice", "--headless", "--convert-to", "pdf",
    "input.pptx", "--outdir", "output/"
])

# PDF 转 PNG
doc = fitz.open("output.pdf")
for page_num in range(len(doc)):
    page = doc[page_num]
    pix = page.get_pixmap()
    pix.save(f"output/{page_num + 1}.png")
```

## 命令行用法

```bash
# 基本转换
python3 convert.py --input /path/to/file.pptx --output ./output

# 带 ZIP 打包
python3 convert.py --input /path/to/file.pptx --output ./output --zip

# 指定输出文件夹
python3 convert.py -i presentation.pptx -o ./images
```

## 📖 文档

- [English README](README.md)
- [中文文档](README-CN.md)
- [技能定义](SKILL.md)

## 🔧 环境要求

- Python 3.8+
- LibreOffice（无头模式）
- PyMuPDF (fitz)

## 📂 项目结构

```
ppt-to-png/
├── SKILL.md           # OpenClaw 技能定义
├── convert.py         # 主 Python 模块
├── README.md          # 英文文档
├── README-CN.md       # 中文文档
├── LICENSE            # MIT 许可协议
└── .gitignore
```

## 🤝 贡献

欢迎提交 Pull Request！

## 📜 许可证

MIT License - 见 [LICENSE](LICENSE) 文件。

## 👥 作者

- **魏然 (Weiran)** - [GitHub](https://github.com/QiaoTuCodes)
- **焱焱 (Yanyan)** - yanyan@3c3d77679723a2fe95d3faf9d2c2e5a65559acbc97fef1ef37783514a80ae453

## 🙏 致谢

- [LibreOffice](https://www.libreoffice.org/) - 开源办公套件
- [PyMuPDF](https://pymupdf.readthedocs.io/) - Python PDF 处理库
- [OpenClaw](https://github.com/openclaw/openclaw) 团队

---

<p align="center">
  <sub>用 ❤️ 为 OpenClaw 社区打造</sub>
</p>
