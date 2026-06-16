# 公式转换工具

将 Word 文档中的纯文本数学公式自动识别并替换为 OMML 公式，保留原文档样式。

---

## 功能

- 基于 CJK 边界 + 数学特征自动识别公式
- 纯文本 → LaTeX → MathML → OMML 全自动转换
- 保留原格式（字体、缩进、行距）
- 支持 .doc / .docx 输入
- 通过 Word/WPS COM 自动修正公式字体
- 转换失败自动跳过，不卡住
- GUI / 命令行双模式

## 快速开始

**安装**（需要 Python 3.10+）

```
双击 install.bat
```

**运行**

```
.venv\Scripts\python.exe main.py
```

**打包为 exe**

```
.venv\Scripts\python.exe build.py
```

生成 `dist/公式转换工具.exe`，单文件免安装运行。

## 命令行

```
.venv\Scripts\python.exe src\replace_formulas.py 输入文件.docx
.venv\Scripts\python.exe src\replace_formulas.py 输入文件.docx -o 输出文件.docx
```

不指定 `-o` 时默认输出 `输入文件名_公式版.docx`。

## 识别规则

| 规则 | 示例 |
|------|------|
| 数学符号 | `α+β`, `∫f(x)dx`, `∞` |
| 上下标 | `x²`, `a₀`, `I_norm` |
| 分数 | `a/b`, `ΔS/T` |
| 关系符+变量 | `F=ma`, `p<0.05`, `PV=nRT` |
| 算术表达式 | `2x+3y-5`, `a+b=c` |
| 函数调用 | `f(x)=0`, `cos(θ)=1` |

**自动排除**

| 类型 | 示例 |
|------|------|
| 物理单位 | `km/s`, `N·m`, `μmol/L`, `s⁻¹` |
| 单位赋值 | `λ=589nm`, `v₀=5m/s` |
| 简单阈值 | `≥100`, `=3.5` |
| 英文文本 | `the value > 0`, `and = 5` |

## 项目结构

```
main.py                       入口
build.py                      打包
install.bat                   安装依赖
requirements.txt              依赖列表
assets/
  icon.ico / icon.svg         图标
src/
  formula_detector.py         公式识别 + LaTeX 转换
  replace_formulas.py         文档处理核心
  mathml_to_omml.py           MathML → OMML
  formula_converter_gui.py    GUI
```

## 依赖

| 包 | 用途 |
|---|------|
| python-docx | Word 文档读写 |
| lxml | XML 处理 |
| latex2mathml | LaTeX → MathML |
| pyinstaller | 打包为 exe |
