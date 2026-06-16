# -*- coding: utf-8 -*-
"""
公式识别器：从段落文本中识别数学公式片段，并转为 LaTeX。

核心策略：中文字符（CJK）是天然的公式边界。
公式主体不含中文字符，由中文文本/标点界定。
"""
import re

GREEK_LETTERS = 'αβγδεζηθλμξπρσφψωϑϵϱϰϖϕΑΒΓΔΕΖΗΘΛΜΞΠΡΣΦΨΩϒ'
MATH_SYMBOLS = '²³⁰¹⁴⁵⁶⁷⁸⁹ⁿΣ∏∫√∂∇∞≈≠≤≥±×÷·→←↑↓∓∮∬∭∛∜∈∉⊂⊃⊆⊇∅∩∪∀∃∧∨¬∴∵∠⊥∥⊕⊗⊙⊖⊘∝≡≅≪≫≺≻≲≳≾≿≁≂⋆∗∘•ℏℓℵ⇌⇀⇁↼↽⊞⊟⊠⊡⨀⨁⨂†‡'
SUPERSCRIPTS = str.maketrans('²³⁰¹⁴⁵⁶⁷⁸⁹ⁿ', '2301456789n')
SUBSCRIPT_MAP = str.maketrans('₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ', '0123456789aehijklmnoprstuvx')

GREEK_TO_LATEX = {
    # 小写希腊字母
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta',
    'ε': r'\varepsilon', 'ϵ': r'\epsilon', 'ζ': r'\zeta', 'η': r'\eta',
    'θ': r'\theta', 'ϑ': r'\vartheta', 'λ': r'\lambda', 'μ': r'\mu',
    'ξ': r'\xi', 'π': r'\pi', 'ϖ': r'\varpi', 'ρ': r'\rho', 'ϱ': r'\varrho',
    'σ': r'\sigma', 'φ': r'\varphi', 'ϕ': r'\phi', 'ψ': r'\psi',
    'ω': r'\omega', 'ϰ': r'\varkappa',
    # 大写希腊字母
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Ε': r'\Epsilon', 'Ζ': r'\Zeta',
    'Η': r'\Eta', 'Θ': r'\Theta', 'Κ': r'\Kappa', 'Λ': r'\Lambda',
    'Ξ': r'\Xi', 'Π': r'\Pi', 'Σ': r'\Sigma', 'Φ': r'\Phi',
    'Ψ': r'\Psi', 'Ω': r'\Omega', 'ϒ': r'\Upsilon',
    # 数学算子
    '∂': r'\partial', '∇': r'\nabla', '∞': r'\infty',
    # 关系运算符
    '≈': r'\approx', '≠': r'\neq', '≤': r'\leq', '≥': r'\geq',
    '∓': r'\mp', '∝': r'\propto', '≡': r'\equiv', '≅': r'\cong',
    '≪': r'\ll', '≫': r'\gg',
    '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊃': r'\supset',
    '⊆': r'\subseteq', '⊇': r'\supseteq', '∅': r'\emptyset',
    '∩': r'\cap', '∪': r'\cup',
    '≺': r'\prec', '≻': r'\succ',
    '≲': r'\lesssim', '≳': r'\gtrsim',
    '≾': r'\precsim', '≿': r'\succsim',
    '≁': r'\nsim', '≂': r'\eqsim',
    # 逻辑
    '∀': r'\forall', '∃': r'\exists',
    '∧': r'\land', '∨': r'\lor', '¬': r'\lnot',
    '∴': r'\therefore', '∵': r'\because',
    # 几何
    '∠': r'\angle', '⊥': r'\perp', '∥': r'\parallel',
    # 四则运算
    '±': r'\pm', '×': r'\times', '÷': r'\div', '·': r'\cdot',
    '⋅': r'\cdot', '∗': r'\ast', '⋆': r'\star', '∘': r'\circ',
    '•': r'\bullet', '∖': r'\setminus',
    '⊕': r'\oplus', '⊗': r'\otimes', '⊙': r'\odot', '⊖': r'\ominus', '⊘': r'\oslash',
    '⊞': r'\boxplus', '⊟': r'\boxminus', '⊠': r'\boxtimes', '⊡': r'\boxdot',
    # 积分
    '∮': r'\oint', '∬': r'\iint', '∭': r'\iiint',
    '⨀': r'\bigodot', '⨁': r'\bigoplus', '⨂': r'\bigotimes',
    # 根号
    '√': r'\sqrt', '∛': r'\sqrt[3]', '∜': r'\sqrt[4]',
    # 箭头
    '→': r'\rightarrow', '←': r'\leftarrow', '↑': r'\uparrow', '↓': r'\downarrow',
    '↔': r'\leftrightarrow', '⇒': r'\Rightarrow', '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow',
    '⇀': r'\rightharpoonup', '⇁': r'\rightharpoondown',
    '↼': r'\leftharpoonup', '↽': r'\leftharpoondown',
    '⇌': r'\rightleftharpoons',
    # 特殊符号
    '…': r'\ldots', '′': r"'", '″': r"''", '°': r'^\circ',
    '†': r'\dagger', '‡': r'\ddagger',
    'ℏ': r'\hbar', 'ℓ': r'\ell', 'ℵ': r'\aleph', 'ℶ': r'\beth',
    '≀': r'\wr',
}

# 中文字符正则（CJK Unified Ideographs + 常用扩展）
CN_CHAR_RE = re.compile(r'[一-鿿㐀-䶿]')
# 中文标点
CN_PUNCT = '，。；！？：、'
CN_BRACKETS_OPEN = '（'
CN_BRACKETS_CLOSE = '）'
# 关系运算符
RELATION_OPS = '=≈≠≤≥<>≡∝≢≅≪≫∈∉⊂⊃⊆⊇≺≻≲≳≾≿≁≂'
# 数学运算符字符（非中文、非纯字母数字的数学专用字符）
MATH_OP_CHARS = set(
    # 上下标
    '²³⁰¹⁴⁵⁶⁷⁸⁹ⁿ₀₁₂₃₄₅₆₇₈₉'
    # 希腊字母
    'αβγδεζηθλμξπρσφψωϑϵϱϰϖϕΑΒΓΔΕΖΗΘΛΜΞΠΡΣΦΨΩϒ'
    # 算子
    '√∛∜∂∇∞'
    # 四则运算
    '±∓×÷·⋅∗⋆∘•∖'
    # 关系
    '≈≠≤≥<>≡∝≢≅≪≫∈∉⊂⊃⊆⊇≺≻≲≳≾≿≁≂'
    # 逻辑
    '∀∃∧∨¬∴∵'
    # 几何
    '∠⊥∥'
    # 集合
    '∅∩∪'
    # 圈运算
    '⊕⊗⊙⊖⊘⊞⊟⊠⊡'
    # 大运算
    'Σ∏∫∮∬∭⨀⨁⨂'
    # 箭头
    '→←↑↓↔⇒⇐⇔⇀⇁↼↽⇌'
    # 其他
    '…′″°†‡ℏℓℵ≀'
)


def has_formula_features(text):
    """判断文本是否包含公式特征（用于快速筛选，粗筛即可，细判交给 _is_valid_formula）"""
    has_math_sym = any(c in MATH_SYMBOLS + GREEK_LETTERS for c in text)
    has_frac = '/' in text and re.search(r'[A-Za-z0-9)]/[A-Za-z0-9(]', text) is not None
    has_super = any(c in '²³⁰¹⁴⁵⁶⁷⁸⁹ⁿ' for c in text)
    has_sub = any(c in '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ' for c in text)
    has_relation = any(c in RELATION_OPS for c in text)
    has_underscore_var = '_' in text and re.search(r'[A-Za-z]+_[A-Za-z0-9]+', text) is not None
    has_letter = bool(re.search(r'[A-Za-z]', text))
    has_number = bool(re.search(r'\d', text))

    if has_math_sym or has_super or has_sub or has_frac:
        return True
    # 有关系符+变量特征（粗筛：关系符+字母+数字共存或函数调用）
    if has_relation and (has_underscore_var or re.search(r'[A-Za-z]\(', text)
                        or (has_letter and has_number)
                        or re.search(r'[A-Za-z]{1,2}[=≈≠≤≥<>]', text)
                        or re.search(r'[=≈≠≤≥<>][A-Za-z]{1,2}', text)):
        return True
    return False


def _is_unit_text(text):
    """判断是否为物理单位（如 km/km², mg/m³, °C/m）或单位赋值（如 u=3.5m/s）而非公式"""
    unit_patterns = [
        r'^[a-zA-Z]+/[a-zA-Z]+[²³]?$',
        r'^[°]?[A-Za-z]+/[A-Za-z]+[²³]?$',
        r'^[A-Za-z]+/[A-Za-z]+$',
        r'^[A-Za-z]\s*[=≈]\s*\d+\.?\d*\s*[°]?[A-Za-z]+(/[A-Za-z]+[²³]?)?$',
        r'^[A-Za-zα-ωΑ-Ω]\s*[=≈]\s*\d+\.?\d*\s*[A-Za-z]+[²³]?$',  # λ=589nm, v=5m/s
        r'^[A-Za-z][₀₁₂₃₄₅₆₇₈₉]?\s*[=≈]\s*\d+\.?\d*\s*[A-Za-z]+(/[A-Za-z]+[²³]?)?$',  # v₀=5m/s
        r'^[A-Za-z]+·[A-Za-z]+(/[A-Za-z]+[²³]?)?$',        # N·m, kg·m/s²
        r'^μ[A-Za-z]+(/[A-Za-z]+[²³]?)?$',                  # μmol/L, μg/mL
        r'^[A-Za-z]+(·[A-Za-z]+)?[⁻]?[¹²³⁰]?$',            # s⁻¹, m·s⁻¹
    ]
    text = text.strip()
    for pat in unit_patterns:
        if re.match(pat, text):
            return True
    return False


def _is_trivial_comparison(text):
    """判断是否为简单的比较/阈值（如 ≥100）或英文词+关系符+数字（如 and = 5）而非公式"""
    text = text.strip()
    # 只有关系符+数字，无变量
    if re.match(r'^[≥≤><≈≠=]\s*\d+\.?\d*%?$', text):
        return True
    # 英文单词+关系符+数字（如 and = 5, the > 0），不是公式
    if re.match(r'^[a-zA-Z]{3,}\s+[=≈≠≤≥<>]\s*\d+\.?\d*%?$', text):
        return True
    return False


def _is_english_text(text):
    """判断文本是否像英文散文而非公式：含2+个空格分隔的3+字母英文词"""
    tokens = text.split()
    alpha_words = [t for t in tokens if re.match(r'^[a-zA-Z]{3,}$', t)]
    return len(alpha_words) >= 2


def _is_incomplete_formula(text):
    """判断是否为不完整的公式片段"""
    text = text.strip()
    if re.match(r'^[=≈≤≥<>]\s*\S{0,3}$', text):
        return True
    if len(text) <= 4 and not any(c.isalpha() for c in text):
        return True
    return False


def _is_formula_char(c):
    """判断一个字符是否可能属于公式（非中文、非中文标点）"""
    if CN_CHAR_RE.match(c):
        return False
    if c in CN_PUNCT + CN_BRACKETS_OPEN + CN_BRACKETS_CLOSE:
        return False
    return True


def _find_formula_segments(text):
    """
    在文本中找到所有不含中文字符的连续片段（公式候选）。
    以中文字符和中文标点为边界分割文本。
    返回列表：[(start, end, segment_text), ...]
    """
    segments = []
    i = 0
    n = len(text)
    while i < n:
        if _is_formula_char(text[i]):
            start = i
            while i < n and _is_formula_char(text[i]):
                i += 1
            seg = text[start:i]
            seg_stripped = seg.strip(' \t\n,;')
            if seg_stripped:
                lstrip = len(seg) - len(seg.lstrip(' \t\n,;'))
                segments.append((start + lstrip, start + lstrip + len(seg_stripped), seg_stripped))
        else:
            i += 1
    return segments


def _is_valid_formula(text):
    """判断一个非中文片段是否为有效公式（而非普通英文/数字文本）"""
    text = text.strip()
    if not text or len(text) < 2:
        return False
    if _is_unit_text(text):
        return False
    if _is_incomplete_formula(text):
        return False

    has_relation = any(c in RELATION_OPS for c in text)
    has_math_sym = any(c in MATH_OP_CHARS for c in text)
    has_frac = '/' in text and re.search(r'[A-Za-z0-9)]/[A-Za-z0-9(]', text) is not None
    has_super = any(c in '²³⁰¹⁴⁵⁶⁷⁸⁹ⁿ' for c in text)
    has_sub = any(c in '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ' for c in text)
    has_underscore_var = '_' in text and re.search(r'[A-Za-z]+_[A-Za-z0-9]+', text) is not None
    has_var_number_mix = (re.search(r'(?<![A-Za-z])[A-Za-z]{1,2}\s*[=≈≠≤≥<>≡∝≢≅≪≫∈∉⊂⊃⊆⊇≺≻≲≳≾≿≁≂]\s*\d', text) is not None
                         or re.search(r'\d\s*[=≈≠≤≥<>≡∝≢≅≪≫∈∉⊂⊃⊆⊇≺≻≲≳≾≿≁≂]\s*[A-Za-z]{1,2}(?![A-Za-z])', text) is not None)
    has_number = re.search(r'\d', text) is not None
    has_arith = ('+' in text or '-' in text) and re.search(r'[A-Za-z]', text) is not None and has_number

    # 只要有数学特征就认为是公式候选，但排除简单阈值和英文散文
    if (has_math_sym or has_super or has_sub or has_frac):
        if _is_trivial_comparison(text) or _is_english_text(text):
            return False
        return True
    # 有关系符+下划线变量名（如 I_norm≥0.77）
    if has_relation and has_underscore_var:
        return True
    # 有关系符+函数调用模式（如 f(x)=..., pH=-log[H+]）
    if has_relation and re.search(r'[A-Za-z]\(', text):
        return True
    # 有关系符+变量与数字混合（如 p<0.05, R²≥0.95）
    if has_relation and has_var_number_mix:
        return True
    # 算术表达式（如 2x+3y-5, a+b=c）
    if has_arith:
        return True
    # 有关系符+多变量乘积（如 F=ma, PV=nRT）
    if has_relation and (re.search(r'[A-Za-z]{1,2}[=≈≠≤≥<>]', text) or re.search(r'[=≈≠≤≥<>][A-Za-z]{1,2}', text)) and not _is_english_text(text):
        return True
    # 有关系符+数字，但排除简单阈值和英文散文
    if has_relation and has_number:
        if _is_trivial_comparison(text) or _is_english_text(text):
            return False
        return True

    return False


def _try_extend_formula(text, seg_start, seg_end, seg_text):
    """
    尝试向左和向右扩展公式片段，包含紧邻的英文变量名。
    不会将中文字符拉入公式。
    """
    new_start, new_end = seg_start, seg_end
    extended_text = seg_text

    # 向左扩展：关系符前的变量（如 =0.05 → p=0.05）
    if seg_text and seg_text[0] in RELATION_OPS and seg_start > 0:
        left = seg_start
        while left > 0:
            c = text[left - 1]
            if (c.isascii() and c.isalpha()) or c in '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ_' or c == '|':
                left -= 1
            else:
                break
        if left < seg_start:
            candidate = text[left:new_end]
            if not CN_CHAR_RE.search(candidate) and _is_valid_formula(candidate):
                new_start = left
                extended_text = candidate

    # 向右扩展：数字/关系符后的变量（如 =0.05 → =0.05p）
    if extended_text and extended_text[-1] in '0123456789)%' + ''.join(c for c in RELATION_OPS) and seg_end < len(text):
        right = seg_end
        while right < len(text):
            c = text[right]
            if (c.isascii() and c.isalpha()) or c in '₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ_':
                right += 1
            else:
                break
        if right > seg_end:
            candidate = text[new_start:right]
            if not CN_CHAR_RE.search(candidate) and _is_valid_formula(candidate):
                new_end = right
                extended_text = candidate

    if new_start != seg_start or new_end != seg_end:
        return new_start, new_end, extended_text
    return seg_start, seg_end, seg_text


def _merge_adjacent_formulas(formulas, text):
    """合并相邻的公式片段（中间只有空格或英文逗号分隔）"""
    if len(formulas) <= 1:
        return formulas

    merged = [formulas[0]]
    for f in formulas[1:]:
        prev = merged[-1]
        gap = text[prev[1]:f[0]].strip()
        if not gap or gap in (',', ';', ', ', '; '):
            new_start = prev[0]
            new_end = f[1]
            new_text = text[new_start:new_end]
            merged[-1] = (new_start, new_end, new_text)
        else:
            merged.append(f)
    return merged


def find_formulas_in_text(text):
    """
    从文本中识别公式片段。
    核心策略：以中文字符为边界，提取非中文片段，判断是否为公式。

    返回列表：[(start, end, formula_text), ...]
    """
    if not text or not has_formula_features(text):
        return []

    segments = _find_formula_segments(text)

    formulas = []
    for seg_start, seg_end, seg_text in segments:
        ext_start, ext_end, ext_text = _try_extend_formula(text, seg_start, seg_end, seg_text)
        if _is_valid_formula(ext_text):
            formulas.append((ext_start, ext_end, ext_text))
            continue
        if _is_valid_formula(seg_text):
            formulas.append((seg_start, seg_end, seg_text))

    formulas = _merge_adjacent_formulas(formulas, text)

    result = []
    for start, end, ftext in formulas:
        ftext = ftext.strip(' ,;')
        if ftext and len(ftext) >= 3:
            result.append((start, end, ftext))

    return result


def text_to_latex(formula_text):
    """将纯文本公式转为 LaTeX"""
    s = formula_text.strip()
    s = s.strip('，。、；：！？' + CN_BRACKETS_OPEN + CN_BRACKETS_CLOSE)
    s = s.strip(' ,;')

    # 0. 中文括号 → 英文括号
    s = s.replace(CN_BRACKETS_OPEN, '(').replace(CN_BRACKETS_CLOSE, ')')
    s = s.replace('、', ',')

    # 1. √(...) 和 √X, ∛/∜ 同理
    s = re.sub(r'√\(([^)]+)\)', r'\\sqrt{\1}', s)
    s = re.sub(r'√([A-Za-z0-9]+)', r'\\sqrt{\1}', s)
    s = re.sub(r'∛\(([^)]+)\)', r'\\sqrt[3]{\1}', s)
    s = re.sub(r'∛([A-Za-z0-9]+)', r'\\sqrt[3]{\1}', s)
    s = re.sub(r'∜\(([^)]+)\)', r'\\sqrt[4]{\1}', s)
    s = re.sub(r'∜([A-Za-z0-9]+)', r'\\sqrt[4]{\1}', s)

    # 2. 希腊字母和数学符号替换
    #    希腊字母后紧跟常见下标字母(x,y,z,i,j,k,n,m)视为下标：σy → \sigma_{y}
    #    其他小写字母保持独立：πu 中 u 是独立变量
    #    算子符号(∂∇∥⊥)不做下标处理
    NO_SUBSCRIPT_CHARS = '∂∇∥⊥'
    SUBSCRIPT_CHARS = 'xyzijknm'
    for char, latex in GREEK_TO_LATEX.items():
        if char in '√∛∜':
            continue
        if char in NO_SUBSCRIPT_CHARS:
            s = s.replace(char, f' {latex} ')
            continue
        # 希腊字母+常见下标字母（单个，后面没有更多小写字母）→ 下标形式
        s = re.sub(re.escape(char) + r'([' + SUBSCRIPT_CHARS + r'])(?![a-z])',
                   lambda m: f' {latex}_{{{m.group(1)}}} ', s)
        # 希腊字母+Unicode下标 → 下标形式
        s = re.sub(re.escape(char) + r'([₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ]+)',
                   lambda m: f' {latex}_{{{m.group(1).translate(SUBSCRIPT_MAP)}}} ', s)
        # 其余希腊字母仅替换符号
        s = s.replace(char, f' {latex} ')

    # 3. 上标：Unicode 上标数字
    def replace_superscripts(m):
        base = m.group(1) if m.group(1) else ''
        sups = m.group(2).translate(SUPERSCRIPTS)
        return f'{base}^{{{sups}}}'
    s = re.sub(r'(\w)\s*([⁰¹²³⁴⁵⁶⁷⁸⁹]+)', replace_superscripts, s)
    s = re.sub(r'(?<!\w)([⁰¹²³⁴⁵⁶⁷⁸⁹]+)', lambda m: '^{' + m.group(1).translate(SUPERSCRIPTS) + '}', s)

    # 4. 下标：Unicode 下标
    def replace_subscripts(m):
        base = m.group(1)
        subs = m.group(2).translate(SUBSCRIPT_MAP)
        return f'{base}_{{{subs}}}'
    s = re.sub(r'(\w)\s*([₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ]+)', replace_subscripts, s)

    # 5. 下划线下标：x_i → x_{i}
    s = re.sub(r'(\w)_([A-Za-z0-9]+)', r'\1_{\2}', s)

    # 6. 分数（先清理 / 前后多余空格，希腊字母替换产生的）
    s = re.sub(r'\s*/\s*', '/', s)
    s = re.sub(r'\(([^)]+)\)/\(([^)]+)\)', r'\\frac{\1}{\2}', s)
    s = re.sub(r'((?:[A-Za-z0-9]+|\\[a-zA-Z]+)(?:\^?\{[^}]+\})?)/((?:[A-Za-z0-9]+|\\[a-zA-Z]+)(?:\^?\{[^}]+\})?)',
               r'\\frac{\1}{\2}', s)

    # 7. exp → e^{...}（支持嵌套括号）
    def replace_exp(m):
        inner = m.group(1)
        return f'e^{{{inner}}}'
    # 用括号配对匹配 exp(...)
    def _find_exp_matches(s):
        result = []
        i = 0
        while i < len(s):
            if s[i:i+4] == 'exp(' and (i == 0 or not s[i-1].isalpha()):
                depth = 0
                j = i + 3  # 指向 '('
                start = j + 1
                while j < len(s):
                    if s[j] == '(':
                        depth += 1
                    elif s[j] == ')':
                        depth -= 1
                        if depth == 0:
                            result.append((i, j + 1, s[start:j]))
                            break
                    j += 1
            i += 1
        return result

    exp_matches = _find_exp_matches(s)
    for start, end, inner in reversed(exp_matches):
        s = s[:start] + f'e^{{{inner}}}' + s[end:]

    # 8. ln → \ln
    s = re.sub(r'\bln\b', r'\\ln', s)

    # 9. 函数名
    for func in ['cos', 'sin', 'tan', 'log', 'max', 'min', 'atan2', 'atan']:
        s = re.sub(r'\b' + func + r'\b', r'\\' + func, s)

    # 10. 清理
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'\\sqrt\s*\{', r'\\sqrt{', s)

    return s
