# 公式转换工具

将 Word 文档中的纯文本数学公式自动识别并转换为 OMML 公式，保留原文档样式。

---

## 功能特性

- **自动识别**：基于 CJK 边界 + 数学特征智能识别公式
- **自动转换**：纯文本 → LaTeX → MathML → OMML 全自动转换
- **保留格式**：保留原文档字体、缩进、行距等样式
- **兼容广泛**：支持 .doc / .docx 格式输入
- **字体修正**：通过 Word/WPS COM 自动修正公式字体
- **容错处理**：转换失败自动跳过，不中断流程
- **双模式支持**：GUI 图形界面 / 命令行两种使用方式

## 快速开始

### 环境要求

- Python 3.10+
- Windows 系统（依赖 Word/WPS COM）

### 安装

双击运行安装脚本：

```
install.bat
```

### 运行

**GUI 模式：**

```
.venv\Scripts\python.exe main.py
```

**命令行模式：**

```
.venv\Scripts\python.exe src\replace_formulas.py 输入文件.docx
.venv\Scripts\python.exe src\replace_formulas.py 输入文件.docx -o 输出文件.docx
```

> 不指定 `-o` 时，默认输出 `输入文件名_公式版.docx`

### 打包为 EXE

```
.venv\Scripts\python.exe build.py
```

生成 `dist/公式转换工具.exe`，单文件免安装运行。

## 识别规则

| 规则 | 示例 |
|------|------|
| 数学符号 | `α+β`, `∫f(x)dx`, `∞` |
| 上下标 | `x²`, `a₀`, `I_norm` |
| 分数 | `a/b`, `ΔS/T` |
| 关系符+变量 | `F=ma`, `p<0.05`, `PV=nRT` |
| 算术表达式 | `2x+3y-5`, `a+b=c` |
| 函数调用 | `f(x)=0`, `cos(θ)=1` |

### 自动排除内容

| 类型 | 示例 |
|------|------|
| 物理单位 | `km/s`, `N·m`, `μmol/L`, `s⁻¹` |
| 单位赋值 | `λ=589nm`, `v₀=5m/s` |
| 简单阈值 | `≥100`, `=3.5` |
| 英文文本 | `the value > 0`, `and = 5` |

## 项目结构

```
formula-converter/
├── main.py                       # 程序入口
├── build.py                      # 打包脚本
├── install.bat                   # 安装依赖
├── requirements.txt              # 依赖列表
├── assets/
│   ├── icon.ico                  # 图标
│   ├── icon.png                  # 图标
│   └── icon.svg                  # 图标
└── src/
    ├── formula_detector.py       # 公式识别 + LaTeX 转换
    ├── replace_formulas.py       # 文档处理核心
    ├── mathml_to_omml.py         # MathML → OMML 转换
    └── formula_converter_gui.py  # GUI 界面
```

## 依赖列表

| 包 | 用途 |
|---|------|
| python-docx | Word 文档读写 |
| lxml | XML 处理 |
| latex2mathml | LaTeX → MathML |
| pyinstaller | 打包为 EXE |

## 开源协议

MIT License