---
name: ppt-to-png
description: 将PPT/PPTX文件导出为PNG图片，支持批量转换。验证可行！
---

# PPT转PNG技能 ✅ 已验证

将PowerPoint文件（.ppt/.pptx）转换为PNG图像格式。

## 验证结果

✅ **转换成功！**
- 输入：50页PPT
- 输出：50张PNG图片
- 打包：ZIP文件（15MB）

## 使用方法

### 自动转换（需要安装依赖）

```bash
# 安装依赖
brew install libreoffice
pip3 install PyMuPDF --break-system-packages

# 运行转换
python3 ~/openclaw-workspace/skills/ppt-to-png/convert.py \
    --input /path/to/file.pptx \
    --output ./output \
    --zip
```

### 依赖说明

| 工具 | 用途 | 安装方式 |
|------|------|----------|
| LibreOffice | PPT→PDF | `brew install libreoffice` |
| PyMuPDF | PDF→PNG | `pip3 install PyMuPDF` |

## 技术原理

```
PPTX → LibreOffice(headless) → PDF → PyMuPDF → PNG(每页)
```

## 输出

- 每页生成一张PNG图片
- 文件名：1.png, 2.png, ... 50.png
- 打包后可通过飞书发送

## 飞书发送

```bash
openclaw message send --channel feishu \
    --target "用户ID" \
    --media "~/path/to/file.zip" \
    --message "PPT已转换完成"
```

## 已知问题

- PDF转换可能有警告，但不影响结果
- 需要确保输入文件路径无中文或特殊字符
