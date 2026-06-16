# WF Conversion Tool

将 Word 文档中的纯文本数学公式自动识别并转换为 OMML 公式，保留原样式。

## 功能

- **自动识别公式**：从段落文本中识别数学公式片段，并转为 LaTeX。
- **保留原样式**：在保留原文档样式的前提下，自动识别段落中的数学公式并替换为 OMML 公式。
- **支持 .doc 转 .docx**：自动将 .doc 文件转为 .docx（使用 Word 或 WPS COM）。
- **GUI 界面**：提供简单的图形用户界面，方便操作。

## 安装

### 环境要求

- [uv](https://docs.astral.sh/uv/) - Python 包管理器（若未安装，参考[官方安装指南](https://docs.astral.sh/uv/getting-started/installation/)）
- Python >= 3.10
- Microsoft Word 或 WPS Office（用于 .doc 转 .docx）

### 初始化项目

```bash
uv init
```

### 构建环境

```bash
uv sync
```

## 使用

### 运行 GUI

```bash
uv run python main.py
```

### 打包为 EXE

```bash
uv run python build.py
```

打包后的可执行文件位于 `dist/WF Conversion Tool.exe`。

## 项目结构

```
formula-converter/
├── main.py                      # 入口文件
├── build.py                     # 打包脚本
├── pyproject.toml               # 项目配置
├── src/
│   ├── formula_converter_gui.py # GUI 主逻辑
│   ├── formula_detector.py      # 公式识别器
│   ├── replace_formulas.py      # 公式替换核心逻辑
│   └── mathml_to_omml.py      # MathML → OMML 转换器
└── assets/
    ├── icon.png
    ├── icon.ico
    └── icon.svg
```

## 核心模块

### formula_detector.py
公式识别器，从段落文本中识别数学公式片段，并转为 LaTeX。核心策略：中文字符（CJK）是天然的公式边界。

### replace_formulas.py
在保留原文档样式的前提下，自动识别段落中的数学公式并替换为 OMML 公式。

### mathml_to_omml.py
将 latex2mathml 生成的 MathML XML 转换为 Word OMML XML。

## 技术栈

- [python-docx](https://python-docx.readthedocs.io/) - 读写 Word 文档
- [lxml](https://lxml.de/) - XML 处理
- [latex2mathml](https://github.com/roniemartinez/latex2mathml) - LaTeX 转 MathML
- [PyInstaller](https://pyinstaller.org/) - 打包为 EXE

## 许可证

MIT
