# -*- coding: utf-8 -*-
"""
在保留原文档样式的前提下，自动识别段落中的数学公式并替换为 OMML 公式
"""
import copy
import gc
import os
import re
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _mark_formula_as_skipped(para_elem, run_start, run_end, formula_text):
    """
    将失败的公式 runs 合并为一个 run，并在文本前插入零宽空格（U+200B），
    使公式检测器不再匹配该文本，同时保留原文内容在文档中的显示。
    """
    runs = para_elem.findall(f'{{{WORD_NS}}}r')
    if run_start >= len(runs):
        return

    # 收集合并文本
    run_texts = []
    for run in runs[run_start:run_end+1]:
        t = run.find(f'{{{WORD_NS}}}t')
        run_texts.append(t.text if t is not None and t.text else '')

    merged_text = ''.join(run_texts)

    # 获取第一个 run 的格式
    first_run = runs[run_start]
    rPr = first_run.find(f'{{{WORD_NS}}}rPr')

    # 创建新 run，文本前加零宽空格
    new_run = etree.Element(f'{{{WORD_NS}}}r')
    if rPr is not None:
        new_run.append(copy.deepcopy(rPr))
    new_t = etree.SubElement(new_run, f'{{{WORD_NS}}}t')
    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    new_t.text = '​' + merged_text  # 零宽空格前缀，破坏公式匹配

    # 替换原 runs
    pos = list(para_elem).index(runs[run_start])
    for i in range(run_start, run_end + 1):
        para_elem.remove(runs[i])
    para_elem.insert(pos, new_run)


def _find_one_formula_in_para(para_elem):
    """
    在段落中识别一个公式片段（找到最后一个，方便从后往前替换）。
    返回 (run_start, run_end, formula_text) 或 None。

    策略：
    1. 获取段落完整文本，用 CJK 边界分割提取公式片段
    2. 将公式文本映射回 run 的位置范围
    """
    from formula_detector import find_formulas_in_text

    runs = para_elem.findall(f'{{{WORD_NS}}}r')
    if not runs:
        return None

    # 收集每个 run 的文本和累计偏移
    run_texts = []
    cum_offsets = [0]
    for run in runs:
        t = run.find(f'{{{WORD_NS}}}t')
        raw = (t.text if t is not None and t.text else '')
        run_texts.append(raw)
        cum_offsets.append(cum_offsets[-1] + len(raw))

    full_text = ''.join(run_texts)
    if not full_text.strip():
        return None

    # 用 CJK 边界提取公式
    text_formulas = find_formulas_in_text(full_text)
    if not text_formulas:
        return None

    # 取最后一个公式（从后往前替换）
    f_start, f_end, f_text = text_formulas[-1]

    # 将文本偏移映射到 run 范围
    r_start = None
    r_end = None
    for i in range(len(runs)):
        r_offset = cum_offsets[i]
        r_len = len(run_texts[i])
        if r_offset + r_len > f_start and r_offset < f_end:
            if r_start is None:
                r_start = i
            r_end = i

    if r_start is not None and r_end is not None:
        return (r_start, r_end, f_text)
    return None


def _replace_runs_with_omml(para_elem, run_start, run_end, formula_text, omml_para_element):
    """
    将段落中 run[run_start..run_end] 替换为 OMML 公式。
    处理公式文本在合并 run 文本中的位置，保留前后文本。
    """
    runs = para_elem.findall(f'{{{WORD_NS}}}r')

    # 收集合并文本
    run_texts = []
    for run in runs[run_start:run_end+1]:
        t = run.find(f'{{{WORD_NS}}}t')
        run_texts.append(t.text if t is not None and t.text else '')

    merged_text = ''.join(run_texts)

    # 在合并文本中找到公式位置
    idx = merged_text.find(formula_text)
    if idx == -1:
        stripped_merged = merged_text.strip()
        stripped_formula = formula_text.strip()
        idx2 = stripped_merged.find(stripped_formula)
        if idx2 != -1:
            lstrip = len(merged_text) - len(merged_text.lstrip())
            idx = lstrip + idx2
        else:
            idx = 0

    before_text = merged_text[:idx]
    after_text = merged_text[idx + len(formula_text):]

    # 获取第一个 run 的格式（直接引用，不 deepcopy）
    first_run = runs[run_start]
    rPr = first_run.find(f'{{{WORD_NS}}}rPr')

    new_elements = []

    if before_text:
        new_run = etree.Element(f'{{{WORD_NS}}}r')
        if rPr is not None:
            new_run.append(copy.deepcopy(rPr))
        new_t = etree.SubElement(new_run, f'{{{WORD_NS}}}t')
        new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_t.text = before_text
        new_elements.append(new_run)

    new_elements.append(omml_para_element)

    if after_text:
        new_run = etree.Element(f'{{{WORD_NS}}}r')
        if rPr is not None:
            new_run.append(copy.deepcopy(rPr))
        new_t = etree.SubElement(new_run, f'{{{WORD_NS}}}t')
        new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_t.text = after_text
        new_elements.append(new_run)

    # 替换 runs
    pos = list(para_elem).index(runs[run_start])
    for i in range(run_start, run_end + 1):
        para_elem.remove(runs[i])
    for j, new_elem in enumerate(new_elements):
        para_elem.insert(pos + j, new_elem)

    return True


def _get_office_com_name():
    """检测可用的 Office COM 组件，返回 COM 名称或 None"""
    import winreg
    for name in ['Word.Application', 'Kwps.Application']:
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name)
            winreg.CloseKey(key)
            return name
        except FileNotFoundError:
            continue
    return None


def _is_word_available():
    """检查 Word COM 是否可用"""
    return _get_office_com_name() is not None


def _fix_formula_fonts_via_com(docx_path):
    """通过 Office COM 选中所有公式域，直接设置字体为 Cambria Math"""
    com_name = _get_office_com_name()
    if com_name is None:
        print("提示：本机未安装 Word 或 WPS，无法自动修正公式字体")
        return

    import subprocess
    import tempfile

    docx_ps_path = docx_path.replace(os.sep, '/')

    ps_file = os.path.join(tempfile.gettempdir(), 'fix_fonts.ps1')
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write('try {\n')
        f.write('  $app = New-Object -ComObject ' + com_name + '\n')
        f.write('  $app.Visible = $false\n')
        f.write('  $doc = $app.Documents.Open("' + docx_ps_path + '")\n')
        f.write('  $mathZones = $doc.OMaths\n')
        f.write('  if ($mathZones) {\n')
        f.write('    foreach ($math in $mathZones) {\n')
        f.write('      try {\n')
        f.write('        $math.Font.Name = "Cambria Math"\n')
        f.write('        $math.Font.NameAscii = "Cambria Math"\n')
        f.write('        $math.Font.NameOther = "Cambria Math"\n')
        f.write('      } catch {}\n')
        f.write('    }\n')
        f.write('  }\n')
        f.write('  $doc.Save()\n')
        f.write('  $doc.Close($false)\n')
        f.write('  $app.Quit()\n')
        f.write('  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null\n')
        f.write('  Write-Output "SUCCESS"\n')
        f.write('} catch {\n')
        f.write('  Write-Output "ERROR: $_"\n')
        f.write('}\n')

    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    result = subprocess.run(
        ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
        capture_output=True, timeout=120, startupinfo=si
    )

    if not os.path.exists(docx_path):
        raise RuntimeError("Word COM 字体修正后文件丢失")


def process_document(input_path, output_path, log_func=None):
    from formula_detector import text_to_latex
    from mathml_to_omml import latex_to_omml

    if log_func is None:
        log_func = print

    BATCH_SIZE = 50  # 每处理 50 个公式保存一次

    # 先复制输入文件到输出路径，后续分批在输出文件上操作
    import shutil
    shutil.copy2(input_path, output_path)

    total_replaced = 0
    total_failed = 0
    batch_num = 0
    failed_formulas = set()  # 记录已失败的公式文本，跳过不再重试

    # 分批循环：每批加载文档，处理最多 BATCH_SIZE 个公式，保存，释放
    while True:
        batch_num += 1
        doc = Document(output_path)

        replaced_count = 0
        failed_count = 0

        for para in doc.paragraphs:
            if replaced_count >= BATCH_SIZE:
                break

            para_elem = para._element
            para_replace_limit = 20
            while para_replace_limit > 0 and replaced_count < BATCH_SIZE:
                para_replace_limit -= 1
                result = _find_one_formula_in_para(para_elem)
                if result is None:
                    break

                run_start, run_end, formula_text = result

                # 跳过之前已失败的公式
                if formula_text in failed_formulas:
                    # 标记该 run 范围为已处理，避免无限循环
                    _mark_formula_as_skipped(para_elem, run_start, run_end, formula_text)
                    continue

                try:
                    latex = text_to_latex(formula_text)
                    omml_para = latex_to_omml(latex)
                except Exception as ex:
                    log_func(f'  转换失败: {formula_text[:50]}... ({ex})')
                    failed_count += 1
                    total_failed += 1
                    failed_formulas.add(formula_text)
                    _mark_formula_as_skipped(para_elem, run_start, run_end, formula_text)
                    continue

                if _replace_runs_with_omml(para_elem, run_start, run_end, formula_text, omml_para):
                    replaced_count += 1
                    total_replaced += 1
                    log_func(f'  替换 ({total_replaced}): {formula_text[:60]}')

        if replaced_count == 0:
            del doc
            gc.collect()
            break

        log_func(f'第 {batch_num} 批：替换 {replaced_count} 处，失败 {failed_count} 处，保存中...')

        _set_formula_fonts(doc)

        doc.save(output_path)
        del doc
        gc.collect()

    log_func(f'公式替换完成：成功 {total_replaced} 处，失败 {total_failed} 处')

    log_func("正在通过 Word 修正公式字体...")
    try:
        _fix_formula_fonts_via_com(output_path)
    except Exception as e:
        log_func(f"提示：Word COM 字体修正失败（{e}），请手动在 Word 中全选公式设置为 Cambria Math")

    log_func(f"完成！替换了 {total_replaced} 处公式，输出文件：{output_path}")


def _split_formula_paragraphs(doc):
    """将公式拆分为独立段落，保留原段落缩进和行距"""
    split_count = 0
    paras_to_split = []
    for para in doc.paragraphs:
        para_elem = para._element
        if para_elem.findall(f'.//{{{MATH_NS}}}oMathPara') or para_elem.findall(f'.//{{{MATH_NS}}}oMath'):
            paras_to_split.append(para_elem)

    for para_elem in paras_to_split:
        pPr = para_elem.find(f'{{{WORD_NS}}}pPr')

        if pPr is not None:
            spacing = pPr.find(f'{{{WORD_NS}}}spacing')
            if spacing is not None:
                spacing.set(f'{{{WORD_NS}}}line', '240')
                if f'{{{WORD_NS}}}lineRule' in spacing.attrib:
                    spacing.set(f'{{{WORD_NS}}}lineRule', 'auto')

        segments = []
        current_text_runs = []

        for child in para_elem:
            if child.tag == f'{{{WORD_NS}}}pPr':
                continue
            if child.tag == f'{{{MATH_NS}}}oMathPara' or child.tag == f'{{{MATH_NS}}}oMath':
                if current_text_runs:
                    segments.append(('text', current_text_runs))
                    current_text_runs = []
                segments.append(('formula', [child]))
            else:
                current_text_runs.append(child)

        if current_text_runs:
            segments.append(('text', current_text_runs))

        if len(segments) <= 1:
            continue

        for child in list(para_elem):
            if child.tag != f'{{{WORD_NS}}}pPr':
                para_elem.remove(child)

        first_type, first_children = segments[0]
        for child in first_children:
            para_elem.append(child)

        for seg_type, seg_children in segments[1:]:
            new_para = etree.Element(f'{{{WORD_NS}}}p')

            if pPr is not None:
                # 复制原段落格式，保留缩进位置
                new_pPr = copy.deepcopy(pPr)
                # 公式段落设为单倍行距
                spacing = new_pPr.find(f'{{{WORD_NS}}}spacing')
                if spacing is not None:
                    spacing.set(f'{{{WORD_NS}}}line', '240')
                    if f'{{{WORD_NS}}}lineRule' in spacing.attrib:
                        spacing.set(f'{{{WORD_NS}}}lineRule', 'auto')
                # 公式段落去掉首行缩进，公式顶格与段落左边界对齐
                ind = new_pPr.find(f'{{{WORD_NS}}}ind')
                if ind is not None:
                    # 删除所有首行缩进相关属性
                    for attr in list(ind.attrib.keys()):
                        local = attr.split('}')[1] if '}' in attr else attr
                        if local in ('firstLine', 'firstLineChars', 'hanging', 'hangingChars'):
                            del ind.attrib[attr]
                    # 如果 ind 元素没有属性了，移除它
                    if not ind.attrib:
                        new_pPr.remove(ind)
                new_para.append(new_pPr)

            for child in seg_children:
                new_para.append(child)

            para_elem.addnext(new_para)
            split_count += 1

    if split_count > 0:
        print(f'  拆分了 {split_count} 个公式段落')


def _set_formula_fonts(doc):
    """设置公式字体（Cambria Math / Times New Roman）"""
    for body_elem in doc.element.body.iter():
        if body_elem.tag == f'{{{MATH_NS}}}r':
            t_elem = body_elem.find(f'{{{MATH_NS}}}t')
            if t_elem is None or not t_elem.text:
                continue

            text = t_elem.text
            is_number = all(c.isdigit() or c in '.=+-' for c in text) and any(c.isdigit() for c in text)

            mrPr = body_elem.find(f'{{{MATH_NS}}}rPr')
            if mrPr is None:
                mrPr = etree.Element(f'{{{MATH_NS}}}rPr')
                for i, child in enumerate(body_elem):
                    if child is t_elem:
                        body_elem.insert(i, mrPr)
                        break

            if is_number:
                nor = mrPr.find(f'{{{MATH_NS}}}nor')
                if nor is None:
                    etree.SubElement(mrPr, f'{{{MATH_NS}}}nor')

                wrPr = mrPr.find(f'{{{WORD_NS}}}rPr')
                if wrPr is None:
                    wrPr = etree.SubElement(mrPr, f'{{{WORD_NS}}}rPr')

                rFonts = wrPr.find(f'{{{WORD_NS}}}rFonts')
                if rFonts is None:
                    rFonts = etree.SubElement(wrPr, f'{{{WORD_NS}}}rFonts')

                rFonts.set(f'{{{WORD_NS}}}ascii', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Times New Roman')
                rFonts.set(f'{{{WORD_NS}}}eastAsia', 'Times New Roman')
                rFonts.set(qn('w:cs'), 'Times New Roman')
            else:
                wrPr = mrPr.find(f'{{{WORD_NS}}}rPr')
                if wrPr is None:
                    wrPr = etree.SubElement(mrPr, f'{{{WORD_NS}}}rPr')

                rFonts = wrPr.find(f'{{{WORD_NS}}}rFonts')
                if rFonts is None:
                    rFonts = etree.SubElement(wrPr, f'{{{WORD_NS}}}rFonts')

                rFonts.set(f'{{{WORD_NS}}}ascii', 'Cambria Math')
                rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Cambria Math')
                rFonts.set(f'{{{WORD_NS}}}eastAsia', 'Cambria Math')
                rFonts.set(qn('w:cs'), 'Cambria Math')

    # 添加默认数学字体到文档设置
    settings_elem = doc.settings.element
    mathPr = settings_elem.find(f'{{{WORD_NS}}}mathPr')
    if mathPr is None:
        mathPr = etree.SubElement(settings_elem, f'{{{WORD_NS}}}mathPr')
    mathFont = mathPr.find(f'{{{WORD_NS}}}mathFont')
    if mathFont is None:
        mathFont = etree.SubElement(mathPr, f'{{{WORD_NS}}}mathFont')
    mathFont.set(f'{{{WORD_NS}}}val', 'Cambria Math')


def main():
    import argparse
    parser = argparse.ArgumentParser(description='自动识别Word文档中的数学公式并替换为OMML公式，保留原样式')
    parser.add_argument('input', help='输入文件路径（.doc 或 .docx）')
    parser.add_argument('-o', '--output', help='输出文件路径（默认：输入文件名_公式版.docx）')
    args = parser.parse_args()

    input_path = args.input

    if input_path.lower().endswith('.doc') and not input_path.lower().endswith('.docx'):
        print(f'检测到 .doc 格式，正在转换为 .docx...')
        import subprocess, tempfile, shutil, winreg

        # 检测可用的 Office COM
        com_name = None
        for name in ['Word.Application', 'Kwps.Application']:
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name)
                winreg.CloseKey(key)
                com_name = name
                break
            except FileNotFoundError:
                continue

        if com_name is None:
            print('错误：本机未安装 Microsoft Word 或 WPS Office，无法自动转换 .doc 文件')
            print('请手动将 .doc 另存为 .docx 格式，然后使用 .docx 文件作为输入')
            exit(1)

        app_name = 'Word' if 'Word' in com_name else 'WPS'
        print(f'使用 {app_name} COM 转换...')

        tmp_doc = os.path.join(tempfile.gettempdir(), 'input_temp.doc')
        tmp_docx = os.path.join(tempfile.gettempdir(), 'input_temp.docx')
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)
        shutil.copy2(input_path, tmp_doc)

        doc_ps_path = tmp_doc.replace(os.sep, '/')
        docx_ps_path = tmp_docx.replace(os.sep, '/')

        ps_file = os.path.join(tempfile.gettempdir(), 'convert.ps1')
        with open(ps_file, 'w', encoding='utf-8') as f:
            f.write('try {\n')
            f.write('  $app = New-Object -ComObject ' + com_name + '\n')
            f.write('  $app.Visible = $false\n')
            f.write('  $doc = $app.Documents.Open("' + doc_ps_path + '")\n')
            f.write('  $doc.SaveAs2("' + docx_ps_path + '", 16)\n')
            f.write('  $doc.Close($false)\n')
            f.write('  $app.Quit()\n')
            f.write('  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null\n')
            f.write('  Write-Output "SUCCESS"\n')
            f.write('} catch {\n')
            f.write('  Write-Output "ERROR: $_"\n')
            f.write('}\n')

        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
            capture_output=True, timeout=120,
            startupinfo=subprocess.STARTUPINFO(
                dwFlags=subprocess.STARTF_USESHOWWINDOW,
                wShowWindow=0)
        )
        stdout = result.stdout.decode('utf-8', errors='replace').strip()
        if not os.path.exists(tmp_docx):
            if 'ERROR' in stdout:
                print(f'Word COM 转换失败：{stdout}')
            print('请手动用 Word 将 .doc 另存为 .docx 后再运行。')
            exit(1)

        input_path = tmp_docx
        print(f'转换完成')

    if args.output:
        output_path = args.output
    else:
        base = os.path.splitext(args.input)[0]
        output_path = f'{base}_公式版.docx'

    process_document(input_path, output_path)


if __name__ == '__main__':
    main()
