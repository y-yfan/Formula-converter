# -*- coding: utf-8 -*-
"""
MathML → OMML 转换器
将 latex2mathml 生成的 MathML XML 转换为 Word OMML XML。
复用 replace_formulas.py 中的 make_omml_* 系列函数。
"""
import copy
from lxml import etree
from docx.oxml.ns import qn

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MML_NS = 'http://www.w3.org/1998/Math/MathML'


def _tag(elem):
    """获取元素的本地标签名（去掉命名空间）"""
    tag = elem.tag
    if '}' in tag:
        return tag.split('}')[1]
    return tag


def _text(elem):
    """获取元素文本内容"""
    return (elem.text or '').strip()


def _children(elem):
    """获取非空子元素"""
    return [c for c in elem if isinstance(c.tag, str)]


def _make_omml_run(text):
    """创建 OMML 文本 run：m:r > m:t"""
    r = etree.Element(f'{{{MATH_NS}}}r')
    t = etree.SubElement(r, f'{{{MATH_NS}}}t')
    t.text = text
    return r


def _convert_mi(elem):
    """MathML mi（标识符/变量）→ OMML m:r"""
    text = _text(elem)
    if not text:
        return None
    return _make_omml_run(text)


def _convert_mn(elem):
    """MathML mn（数字）→ OMML m:r"""
    text = _text(elem)
    if not text:
        return None
    return _make_omml_run(text)


def _convert_mo(elem):
    """MathML mo（运算符）→ OMML m:r"""
    text = _text(elem)
    if not text:
        return None
    # 映射常见运算符
    op_map = {
        '×': '×',   # ×
        '−': '-',   # −
        '⋅': '·',   # ⋅
        '≤': '≤',
        '≥': '≥',
        '≠': '≠',
        '→': '→',
        '←': '←',
        '≈': '≈',
    }
    text = op_map.get(text, text)
    return _make_omml_run(text)


def _convert_mtext(elem):
    """MathML mtext（普通文本）→ OMML m:r"""
    text = _text(elem)
    if not text:
        return None
    return _make_omml_run(text)


def _convert_mspace(elem):
    """MathML mspace（空格）→ 忽略"""
    return None


def _convert_mfrac(elem):
    """MathML mfrac → OMML m:f（分数）"""
    children = _children(elem)
    if len(children) < 2:
        return _convert_group(children)

    f = etree.Element(f'{{{MATH_NS}}}f')
    fPr = etree.SubElement(f, f'{{{MATH_NS}}}fPr')
    num = etree.SubElement(f, f'{{{MATH_NS}}}num')
    den = etree.SubElement(f, f'{{{MATH_NS}}}den')

    # 分子
    num_result = convert(children[0])
    if num_result is not None:
        if isinstance(num_result, list):
            for item in num_result:
                num.append(item)
        else:
            num.append(num_result)

    # 分母
    den_result = convert(children[1])
    if den_result is not None:
        if isinstance(den_result, list):
            for item in den_result:
                den.append(item)
        else:
            den.append(den_result)

    return f


def _convert_msup(elem):
    """MathML msup → OMML m:sSup（上标）"""
    children = _children(elem)
    if len(children) < 2:
        return _convert_group(children)

    sSup = etree.Element(f'{{{MATH_NS}}}sSup')
    e = etree.SubElement(sSup, f'{{{MATH_NS}}}e')
    sup = etree.SubElement(sSup, f'{{{MATH_NS}}}sup')

    base_result = convert(children[0])
    if base_result is not None:
        if isinstance(base_result, list):
            for item in base_result:
                e.append(item)
        else:
            e.append(base_result)

    sup_result = convert(children[1])
    if sup_result is not None:
        if isinstance(sup_result, list):
            for item in sup_result:
                sup.append(item)
        else:
            sup.append(sup_result)

    return sSup


def _convert_msub(elem):
    """MathML msub → OMML m:sSub（下标）"""
    children = _children(elem)
    if len(children) < 2:
        return _convert_group(children)

    sSub = etree.Element(f'{{{MATH_NS}}}sSub')
    e = etree.SubElement(sSub, f'{{{MATH_NS}}}e')
    sub = etree.SubElement(sSub, f'{{{MATH_NS}}}sub')

    base_result = convert(children[0])
    if base_result is not None:
        if isinstance(base_result, list):
            for item in base_result:
                e.append(item)
        else:
            e.append(base_result)

    sub_result = convert(children[1])
    if sub_result is not None:
        if isinstance(sub_result, list):
            for item in sub_result:
                sub.append(item)
        else:
            sub.append(sub_result)

    return sSub


def _convert_msubsup(elem):
    """MathML msubsup → OMML nary（求和等）或 sSubSup（上下标）"""
    children = _children(elem)
    if len(children) < 3:
        return _convert_msub(elem)

    # 检查第一个子元素是否为 n-ary 运算符
    first = children[0]
    if _tag(first) == 'mo':
        text = _text(first)
        nary_map = {
            '∑': '∑', 'Σ': '∑',
            '∫': '∫',
            '∏': '∏',
            '⋃': '⋃', '∪': '⋃',
            '⋂': '⋂', '∩': '⋂',
            '∬': '∬',
            '∮': '∮',
        }
        if text in nary_map:
            nary = etree.Element(f'{{{MATH_NS}}}nary')
            naryPr = etree.SubElement(nary, f'{{{MATH_NS}}}naryPr')
            chr_elem = etree.SubElement(naryPr, f'{{{MATH_NS}}}chr')
            chr_elem.set(f'{{{MATH_NS}}}val', nary_map[text])
            limLoc = etree.SubElement(naryPr, f'{{{MATH_NS}}}limLoc')
            limLoc.set(f'{{{MATH_NS}}}val', 'undOvr')
            sub = etree.SubElement(nary, f'{{{MATH_NS}}}sub')
            sup = etree.SubElement(nary, f'{{{MATH_NS}}}sup')
            e = etree.SubElement(nary, f'{{{MATH_NS}}}e')

            sub_result = convert(children[1])
            if sub_result is not None:
                if isinstance(sub_result, list):
                    for item in sub_result:
                        sub.append(item)
                else:
                    sub.append(sub_result)

            sup_result = convert(children[2])
            if sup_result is not None:
                if isinstance(sup_result, list):
                    for item in sup_result:
                        sup.append(item)
                else:
                    sup.append(sup_result)

            return nary

    # 普通上下标
    sSubSup = etree.Element(f'{{{MATH_NS}}}sSubSup')
    e = etree.SubElement(sSubSup, f'{{{MATH_NS}}}e')
    sub = etree.SubElement(sSubSup, f'{{{MATH_NS}}}sub')
    sup = etree.SubElement(sSubSup, f'{{{MATH_NS}}}sup')

    base_result = convert(children[0])
    if base_result is not None:
        if isinstance(base_result, list):
            for item in base_result:
                e.append(item)
        else:
            e.append(base_result)

    sub_result = convert(children[1])
    if sub_result is not None:
        if isinstance(sub_result, list):
            for item in sub_result:
                sub.append(item)
        else:
            sub.append(sub_result)

    sup_result = convert(children[2])
    if sup_result is not None:
        if isinstance(sup_result, list):
            for item in sup_result:
                sup.append(item)
        else:
            sup.append(sup_result)

    return sSubSup


def _convert_msqrt(elem):
    """MathML msqrt → OMML m:rad（平方根）"""
    rad = etree.Element(f'{{{MATH_NS}}}rad')
    radPr = etree.SubElement(rad, f'{{{MATH_NS}}}radPr')
    degHide = etree.SubElement(radPr, f'{{{MATH_NS}}}degHide')
    degHide.set(f'{{{MATH_NS}}}val', '1')
    deg = etree.SubElement(rad, f'{{{MATH_NS}}}deg')
    deg.append(_make_omml_run(''))
    e = etree.SubElement(rad, f'{{{MATH_NS}}}e')

    children = _children(elem)
    for child in children:
        result = convert(child)
        if result is not None:
            if isinstance(result, list):
                for item in result:
                    e.append(item)
            else:
                e.append(result)

    return rad


def _convert_mroot(elem):
    """MathML mroot → OMML m:rad（n次根）"""
    children = _children(elem)
    if len(children) < 2:
        return _convert_msqrt(elem)

    rad = etree.Element(f'{{{MATH_NS}}}rad')
    radPr = etree.SubElement(rad, f'{{{MATH_NS}}}radPr')
    deg = etree.SubElement(rad, f'{{{MATH_NS}}}deg')
    e = etree.SubElement(rad, f'{{{MATH_NS}}}e')

    # MathML mroot: 第1个是被开方数，第2个是根指数
    base_result = convert(children[0])
    if base_result is not None:
        if isinstance(base_result, list):
            for item in base_result:
                e.append(item)
        else:
            e.append(base_result)

    deg_result = convert(children[1])
    if deg_result is not None:
        if isinstance(deg_result, list):
            for item in deg_result:
                deg.append(item)
        else:
            deg.append(deg_result)

    return rad


def _convert_mfenced(elem):
    """MathML mfenced → OMML m:d（定界符/括号）"""
    open_chr = elem.get('open', '(')
    close_chr = elem.get('close', ')')

    d = etree.Element(f'{{{MATH_NS}}}d')
    dPr = etree.SubElement(d, f'{{{MATH_NS}}}dPr')
    begChr = etree.SubElement(dPr, f'{{{MATH_NS}}}begChr')
    begChr.set(f'{{{MATH_NS}}}val', open_chr)
    endChr = etree.SubElement(dPr, f'{{{MATH_NS}}}endChr')
    endChr.set(f'{{{MATH_NS}}}val', close_chr)
    e = etree.SubElement(d, f'{{{MATH_NS}}}e')

    for child in _children(elem):
        result = convert(child)
        if result is not None:
            if isinstance(result, list):
                for item in result:
                    e.append(item)
            else:
                e.append(result)

    return d


def _get_nary_char(elem):
    """判断 munderover/munder/mover 的运算符字符，返回 (char, is_nary)"""
    # 检查第一个子元素是否是运算符
    children = _children(elem)
    if not children:
        return None, False

    first = children[0]
    tag = _tag(first)

    if tag == 'mo':
        text = _text(first)
        nary_map = {
            '∑': '∑', 'Σ': '∑',
            '∫': '∫',
            '∏': '∏',
            '⋃': '⋃', '∪': '⋃',
            '⋂': '⋂', '∩': '⋂',
            '∬': '∬',
            '∮': '∮',
        }
        if text in nary_map:
            return nary_map[text], True

    if tag == 'mo':
        text = _text(first)
        acc_map = {
            '̅': '̅',  # overbar
            '̄': '̅',  # combining overline
            '⏜': '⏜',  # overbrace
            '˜': '˜',  # tilde
            '→': '→',
            '̂': '̂',  # hat
            '̇': '̇',  # dot
        }
        if text in acc_map:
            return acc_map[text], False

    return None, False


def _convert_munderover(elem):
    """MathML munderover → OMML nary（求和等）或 acc（修饰）"""
    children = _children(elem)
    char, is_nary = _get_nary_char(elem)

    if is_nary and len(children) >= 3:
        # nary: 运算符 + 下限 + 上限 + 主体
        nary = etree.Element(f'{{{MATH_NS}}}nary')
        naryPr = etree.SubElement(nary, f'{{{MATH_NS}}}naryPr')
        chr_elem = etree.SubElement(naryPr, f'{{{MATH_NS}}}chr')
        chr_elem.set(f'{{{MATH_NS}}}val', char)
        limLoc = etree.SubElement(naryPr, f'{{{MATH_NS}}}limLoc')
        limLoc.set(f'{{{MATH_NS}}}val', 'undOvr')
        sub = etree.SubElement(nary, f'{{{MATH_NS}}}sub')
        sup = etree.SubElement(nary, f'{{{MATH_NS}}}sup')
        e = etree.SubElement(nary, f'{{{MATH_NS}}}e')

        # 下限
        sub_result = convert(children[1])
        if sub_result is not None:
            if isinstance(sub_result, list):
                for item in sub_result:
                    sub.append(item)
            else:
                sub.append(sub_result)

        # 上限
        sup_result = convert(children[2])
        if sup_result is not None:
            if isinstance(sup_result, list):
                for item in sup_result:
                    sup.append(item)
            else:
                sup.append(sup_result)

        # 主体（第4个子元素，如果没有则用空）
        if len(children) >= 4:
            e_result = convert(children[3])
            if e_result is not None:
                if isinstance(e_result, list):
                    for item in e_result:
                        e.append(item)
                else:
                    e.append(e_result)

        return nary

    # 非 nary 运算符，按 group 处理
    return _convert_group(children)


def _convert_munder(elem):
    """MathML munder → nary 或 group"""
    children = _children(elem)
    char, is_nary = _get_nary_char(elem)

    if is_nary and len(children) >= 2:
        nary = etree.Element(f'{{{MATH_NS}}}nary')
        naryPr = etree.SubElement(nary, f'{{{MATH_NS}}}naryPr')
        chr_elem = etree.SubElement(naryPr, f'{{{MATH_NS}}}chr')
        chr_elem.set(f'{{{MATH_NS}}}val', char)
        sub = etree.SubElement(nary, f'{{{MATH_NS}}}sub')
        sup = etree.SubElement(nary, f'{{{MATH_NS}}}sup')
        e = etree.SubElement(nary, f'{{{MATH_NS}}}e')

        sub_result = convert(children[1])
        if sub_result is not None:
            if isinstance(sub_result, list):
                for item in sub_result:
                    sub.append(item)
            else:
                sub.append(sub_result)

        if len(children) >= 3:
            e_result = convert(children[2])
            if e_result is not None:
                if isinstance(e_result, list):
                    for item in e_result:
                        e.append(item)
                else:
                    e.append(e_result)

        return nary

    # 修饰符（如横杠）
    if len(children) >= 2 and not is_nary:
        char_val = char if char else '̅'
        acc = etree.Element(f'{{{MATH_NS}}}acc')
        accPr = etree.SubElement(acc, f'{{{MATH_NS}}}accPr')
        chr_elem = etree.SubElement(accPr, f'{{{MATH_NS}}}chr')
        chr_elem.set(f'{{{MATH_NS}}}val', char_val)
        e = etree.SubElement(acc, f'{{{MATH_NS}}}e')

        e_result = convert(children[1])
        if e_result is not None:
            if isinstance(e_result, list):
                for item in e_result:
                    e.append(item)
            else:
                e.append(e_result)

        return acc

    return _convert_group(children)


def _convert_mover(elem):
    """MathML mover → 修饰符或 group"""
    children = _children(elem)
    char, is_nary = _get_nary_char(elem)

    if is_nary and len(children) >= 2:
        return _convert_munder(elem)

    # 修饰符
    if len(children) >= 2:
        char_val = char if char else '̅'
        acc = etree.Element(f'{{{MATH_NS}}}acc')
        accPr = etree.SubElement(acc, f'{{{MATH_NS}}}accPr')
        chr_elem = etree.SubElement(accPr, f'{{{MATH_NS}}}chr')
        chr_elem.set(f'{{{MATH_NS}}}val', char_val)
        e = etree.SubElement(acc, f'{{{MATH_NS}}}e')

        e_result = convert(children[0])
        if e_result is not None:
            if isinstance(e_result, list):
                for item in e_result:
                    e.append(item)
            else:
                e.append(e_result)

        return acc

    return _convert_group(children)


def _convert_mrow(elem):
    """MathML mrow → 顺序排列子元素"""
    results = []
    for child in _children(elem):
        result = convert(child)
        if result is not None:
            if isinstance(result, list):
                results.extend(result)
            else:
                results.append(result)
    return results if len(results) != 1 else results[0]


def _convert_group(children):
    """将多个子元素转为 OMML 元素列表"""
    results = []
    for child in children:
        result = convert(child)
        if result is not None:
            if isinstance(result, list):
                results.extend(result)
            else:
                results.append(result)
    return results[0] if len(results) == 1 else results if results else None


def convert(elem):
    """将 MathML 元素转为 OMML 元素。返回单个元素或列表。"""
    tag = _tag(elem)

    converters = {
        'mi': _convert_mi,
        'mn': _convert_mn,
        'mo': _convert_mo,
        'mtext': _convert_mtext,
        'mspace': _convert_mspace,
        'mfrac': _convert_mfrac,
        'msup': _convert_msup,
        'msub': _convert_msub,
        'msubsup': _convert_msubsup,
        'msqrt': _convert_msqrt,
        'mroot': _convert_mroot,
        'mfenced': _convert_mfenced,
        'munderover': _convert_munderover,
        'munder': _convert_munder,
        'mover': _convert_mover,
        'mrow': _convert_mrow,
        'math': lambda e: _convert_mrow(e),  # 顶层 math 元素
    }

    converter = converters.get(tag)
    if converter:
        return converter(elem)

    # 未知元素：尝试处理子元素
    children = _children(elem)
    if children:
        return _convert_group(children)

    return None


def mathml_to_omml(mathml_str):
    """
    将 MathML 字符串转为 OMML oMathPara 元素。
    返回 lxml.etree.Element，可直接插入 Word 文档段落。
    """
    # 解析 MathML
    mathml_elem = etree.fromstring(mathml_str.encode('utf-8'))

    # 转换为 OMML
    omml_result = convert(mathml_elem)

    # 包装为 oMathPara > oMath
    omath_para = etree.Element(f'{{{MATH_NS}}}oMathPara')
    omath = etree.SubElement(omath_para, f'{{{MATH_NS}}}oMath')

    if omml_result is not None:
        if isinstance(omml_result, list):
            for item in omml_result:
                omath.append(item)
        else:
            omath.append(omml_result)

    return omath_para


def latex_to_omml(latex_str):
    """
    将 LaTeX 字符串转为 OMML oMathPara 元素。
    完整流程：LaTeX → MathML → OMML。
    """
    from latex2mathml.converter import convert as latex_to_mathml

    mathml_str = latex_to_mathml(latex_str)
    return mathml_to_omml(mathml_str)
