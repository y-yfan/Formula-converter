# -*- coding: utf-8 -*-
"""打包脚本：生成单文件 exe"""
import PyInstaller.__main__
import os

# 查找 latex2mathml 的 unimathsymbols.txt
import latex2mathml
lmml_dir = os.path.dirname(latex2mathml.__file__)
unimath = os.path.join(lmml_dir, 'unimathsymbols.txt')

src = os.path.join('src', '')

args = [
    os.path.join('src', 'formula_converter_gui.py'),
    '--name=WF Conversion Tool',
    '--noconsole',
    '--onefile',
    f'--add-data={src}formula_detector.py{os.pathsep}.',
    f'--add-data={src}replace_formulas.py{os.pathsep}.',
    f'--add-data={src}mathml_to_omml.py{os.pathsep}.',
    f'--add-data={unimath}{os.pathsep}latex2mathml',
    '--noconfirm',
]

# 排除未使用的模块以减小体积
excludes = [
    # lxml: 只用 etree
    'lxml.html', 'lxml.html.builder', 'lxml.html.clean', 'lxml.html.defs',
    'lxml.html.diff', 'lxml.html.formfill', 'lxml.html.html5parser',
    'lxml.html.soupparser', 'lxml.cssselect', 'lxml.objectify',
    'lxml.sax', 'lxml.builder', 'lxml.isoschematron',
    'lxml.ElementInclude', 'lxml.usedoctest', 'lxml.doctestcompare',
    # python-docx: 只用 Document + oxml
    'docx.comments', 'docx.dml', 'docx.drawing', 'docx.image',
    'docx.shape', 'docx.styles', 'docx.table',
    'docx.section', 'docx.settings', 'docx.templates',
    # 其他
    'unittest', 'pydoc', 'doctest',
]
for ex in excludes:
    args.append(f'--exclude-module={ex}')

# UPX 压缩（如果可用）
upx_dir = os.path.join(os.path.dirname(__file__), 'upx')
if os.path.isdir(upx_dir):
    args.append(f'--upx-dir={upx_dir}')

# 如果存在图标文件则添加
icon_file = os.path.join(os.path.dirname(__file__), 'assets', 'icon.ico')
if os.path.exists(icon_file):
    args.append(f'--icon={icon_file}')

PyInstaller.__main__.run(args)
