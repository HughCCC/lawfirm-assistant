#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
互联网法律评估意见书生成器
读取 JSON 格式的报告内容，生成符合格式规范的 .docx 文件。

用法：python3 generate_docx.py <input.json> [output.docx]
"""

import json
import sys
import os
from datetime import datetime
from copy import deepcopy

try:
    from docx import Document
    from docx.shared import Pt, Cm, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("错误：需要 python-docx 库。请运行：pip3 install python-docx")
    sys.exit(1)

# 上标数字符号（用于脚注标记）
FOOTNOTE_MARKERS = '¹²³⁴⁵⁶⁷⁸⁹'



# ============================================================
# 违禁字符扫描
# ============================================================
import re

FORBIDDEN_RULES = [
    # Markdown 语法
    (r'\*\*', 'Markdown 加粗语法 (**)'),
    (r'(?:^|\n)#{1,6}\s', 'Markdown 标题标记 (#)'),
    (r'(?<!\w)\*[^*\n]+\*(?!\w)', 'Markdown 斜体 (*)'),
    (r'(?<!\w)_[^_\n]+_(?!\w)', 'Markdown 斜体 (_)'),
    (r'(?:^|\n)\s*[-*]\s', 'Markdown 无序列表 (- 或 *)'),
    (r'(?:^|\n)\s*\d+\.\s', 'Markdown 有序列表 (1.)'),
    (r'\|.*\|', 'Pipe 字符伪表格'),
    (r'(?:^|\n)[-=*_]{3,}\s*$', '装饰性分割线'),
    (r'`[^`]+`', '行内代码标记 (`)'),
    (r'(?:^|\n)>\s', 'Blockquote 标记 (>)'),
    (r'~~.+?~~', 'Markdown 删除线 (~~)'),
    (r'\[.+\]\(.+\)', 'Markdown 链接语法'),

    # AI 风格标签
    (r'【[^】]*(?:风险|结论|建议|注意|提示|分析|警告)[^】]*】', 'AI 风格方括号标签'),

    # Emoji
    (r'[🟢🟡🟠🔴⚠️📋❌✅🎯💡📌🔍⭐🚫⛔✔️✖️🔥💣🧩🎮🛡️📊📈📉🔒🔑💻📱🖥️🤖👤👥]', 'Emoji 符号'),

    # ASCII 装饰
    (r'[—]{3,}', 'ASCII 装饰线 (——)'),
    (r'[=]{3,}', 'ASCII 等号分隔 (===)'),
    (r'[*]{3,}', 'ASCII 星号分隔 (***)'),
    (r'[#]{3,}', 'ASCII 井号分隔 (###)'),
]


def scan_forbidden(text):
    """扫描文本中的违禁字符，返回违禁项列表。"""
    violations = []
    for pattern, description in FORBIDDEN_RULES:
        matches = re.findall(pattern, text)
        if matches:
            violations.append({
                'rule': description,
                'pattern': pattern,
                'matches': list(set(matches))[:5]  # 最多展示5个
            })
    return violations


def scan_json_content(data):
    """递归扫描 JSON 内容中的所有文本字段。"""
    violations = []
    if isinstance(data, str):
        return scan_forbidden(data)
    elif isinstance(data, dict):
        for key, value in data.items():
            violations.extend(scan_json_content(value))
    elif isinstance(data, list):
        for item in data:
            violations.extend(scan_json_content(item))
    return violations


# ============================================================
# LQ/RQ 兜底替换
# ============================================================

def replace_quote_tokens(text):
    """替换漏网的字面 {LQ}/{RQ} 令牌为实际弯引号字符。

    写报告脚本使用 f-string 中的 {LQ}/{RQ} 变量插入中文弯引号（" "），
    但表格行和遗漏 f 前缀的普通字符串会导致令牌直接出现在 JSON 中。
    此函数作为兜底，确保 .docx 中永远不会出现字面 {LQ}/{RQ}。
    """
    if isinstance(text, str):
        return text.replace('{LQ}', '“').replace('{RQ}', '”')
    return text


# ============================================================
# 文档生成
# ============================================================

# 中文字号 → pt 映射
FONT_SIZES = {
    '小二': Pt(18),
    '四号': Pt(14),
    '小四': Pt(12),
    '五号': Pt(10.5),
    '小五': Pt(9),
}


def set_font(run, font_name='宋体', font_name_ascii='Times New Roman'):
    """设置 run 的中英文字体，默认小四。

    注意：w:hAnsi 有意不设置。强制设 w:hAnsi="Times New Roman" 会导致部分
    Word 渲染器对中文文本使用 TNR（该字体无中文 glyph），出现 tofu/乱码。
    仅设 w:ascii（英文 TNR）和 w:eastAsia（中文宋体），让渲染引擎自然回退。
    """
    run.font.name = font_name_ascii
    run.font.size = FONT_SIZES['小四']
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('w:rPr'), {})
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name_ascii)


def set_paragraph_spacing(paragraph, font_size=Pt(12)):
    """设置段落行距为固定 18 磅，段前段后 0.5 行（6 磅）。"""
    pf = paragraph.paragraph_format
    pf.line_spacing = Pt(18)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)


def add_hyperlink_run(paragraph, text, url, font_size_key='四号'):
    """向段落尾部添加一个可点击的外部超链接 run。

    参数：
        paragraph: 目标段落对象
        text:       超链接显示文本
        url:        目标 URL
        font_size_key: 字号键名（'四号' / '小五' 等）
    """
    part = paragraph.part
    r_id = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True
    )

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Hyperlink 样式（蓝色 + 下划线）
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)

    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0563C1')
    rPr.append(c)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    # 中文字体
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

    # 字号
    sz = OxmlElement('w:sz')
    sz_val = str(int(FONT_SIZES[font_size_key] / Pt(1) * 2))  # half-points
    sz.set(qn('w:val'), sz_val)
    rPr.append(sz)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_title(doc, text):
    """添加文档标题：宋体小二加粗居中。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(para, FONT_SIZES['小二'])
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小二']
    run.bold = True
    return para


def add_heading1(doc, text):
    """添加一级标题：宋体小四加粗两端对齐。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = True
    return para


def add_heading2(doc, text):
    """添加二级标题：宋体小四加粗两端对齐。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = True
    return para


def add_heading3(doc, text):
    """添加三级标题：宋体小四加粗两端对齐，首行缩进两字符。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    para.paragraph_format.first_line_indent = Pt(24)
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = True
    return para


def add_heading4(doc, text):
    """添加四级标题：宋体小四不加粗两端对齐，首行缩进两字符。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    para.paragraph_format.first_line_indent = Pt(24)
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = False
    return para


def add_heading5(doc, text):
    """添加五级标题：宋体小四不加粗两端对齐，首行缩进两字符。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    para.paragraph_format.first_line_indent = Pt(24)
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = False
    return para


def add_body(doc, text):
    """添加正文：宋体小四两端对齐，首行缩进两字符，行距固定 18 磅。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    # 首行缩进两字符（小四 = 12pt，两字符 = 24pt）
    para.paragraph_format.first_line_indent = Pt(24)
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = False
    return para


def add_body_with_line_breaks(doc, text):
    """添加含内部换行的正文：宋体小四左对齐，首行缩进两字符。

    将 \\n 转换为 Word 原生 <w:br/> 换行标签，使用左对齐避免短行被
    两端对齐拉伸（这是第11-14页出现大量空格的根因）。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    para.paragraph_format.first_line_indent = Pt(24)

    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            # 在 run 之间插入 <w:br/> 原生换行标签
            run_br = para.add_run()
            br = OxmlElement('w:br')
            run_br._element.append(br)
        if line.strip():
            run = para.add_run(line)
            set_font(run, '宋体')
            run.font.size = FONT_SIZES['小四']
            run.bold = False

    return para


def add_source_item(doc, text):
    """添加信息来源条目：宋体小四加粗两端对齐，无首行缩进。

    用于"六、信息来源"节中各来源的编号标题（一、二、三…），
    与 add_heading1 视觉一致但语义独立。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(para, FONT_SIZES['小四'])
    run = para.add_run(text)
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['小四']
    run.bold = True
    return para


def add_table(doc, headers, rows, caption=None, note=None, footnote_col=None, hyperlink_col=None, footnote_on_col=None, footnote_urls=None):
    """添加表格：居中，原生 Word 表格，小四字体。

    参数：
        footnote_col: str - 指定某列（匹配 header 名）作为脚注列，该列 URL 将转为
                      上标脚注标记（¹²³…），URL 以超链接脚注形式置于表格下方。
        hyperlink_col: str - 指定某列（匹配 header 名）作为超链接列，该列 URL 将以
                      可点击超链接形式直接嵌入单元格（display 文本为链接本身）。
        footnote_on_col: str - 指定某列（匹配 header 名），在该列文本后追加脚注标记。
                      脚注 URL 从 footnote_urls 数组按行读取。
        footnote_urls: list - 与 rows 一一对应的 URL 数组，供 footnote_on_col 使用。
    """
    # 解析列索引
    f_col_idx = None
    h_col_idx = None
    fon_col_idx = None
    footnotes = []  # [(marker, url), ...]

    if footnote_col:
        for i, h in enumerate(headers):
            if h == footnote_col:
                f_col_idx = i
                break
    if hyperlink_col:
        for i, h in enumerate(headers):
            if h == hyperlink_col:
                h_col_idx = i
                break
    if footnote_on_col:
        for i, h in enumerate(headers):
            if h == footnote_on_col:
                fon_col_idx = i
                break

    # 表格标题（如果有）
    if caption:
        caption = replace_quote_tokens(caption)
        cap_para = doc.add_paragraph()
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(cap_para, FONT_SIZES['小四'])
        run = cap_para.add_run(caption)
        set_font(run, '宋体')
        run.font.size = FONT_SIZES['小四']
        run.bold = True

    # 创建表格
    num_rows = len(rows) + 1  # +1 for header
    num_cols = len(headers)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # 设置表格字体
    for i, header in enumerate(headers):
        header = replace_quote_tokens(header)
        cell = table.rows[0].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].add_run(header)
        set_font(run, '宋体')
        run.font.size = FONT_SIZES['小四']
        run.bold = True
        set_paragraph_spacing(cell.paragraphs[0], FONT_SIZES['小四'])

    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell_text = replace_quote_tokens(str(cell_text)) if cell_text else ''
            cell = table.rows[r + 1].cells[c]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            if f_col_idx is not None and c == f_col_idx:
                # 脚注列：用上标标记替代 URL
                if cell_text and str(cell_text).strip():
                    marker = FOOTNOTE_MARKERS[len(footnotes)] if len(footnotes) < len(FOOTNOTE_MARKERS) else f'[{len(footnotes)+1}]'
                    footnotes.append((marker, str(cell_text).strip()))
                    run = cell.paragraphs[0].add_run(marker)
                    set_font(run, '宋体')
                    run.font.size = FONT_SIZES['小四']
                    run.font.superscript = True
                    run.bold = False
                set_paragraph_spacing(cell.paragraphs[0])
            elif h_col_idx is not None and c == h_col_idx:
                # 超链接列：直接渲染为可点击超链接
                if cell_text and str(cell_text).strip():
                    url = str(cell_text).strip()
                    if url.startswith('http'):
                        add_hyperlink_run(cell.paragraphs[0], url, url, '小五')
                set_paragraph_spacing(cell.paragraphs[0])
            elif fon_col_idx is not None and c == fon_col_idx:
                # footnote_on 列：在原始文本后追加上标脚注标记
                text = str(cell_text) if cell_text else ''
                run = cell.paragraphs[0].add_run(text)
                set_font(run, '宋体')
                run.font.size = FONT_SIZES['小四']
                run.bold = False
                if footnote_urls and r < len(footnote_urls) and footnote_urls[r]:
                    marker = FOOTNOTE_MARKERS[len(footnotes)] if len(footnotes) < len(FOOTNOTE_MARKERS) else f'[{len(footnotes)+1}]'
                    footnotes.append((marker, str(footnote_urls[r]).strip()))
                    run_marker = cell.paragraphs[0].add_run(marker)
                    set_font(run_marker, '宋体')
                    run_marker.font.size = FONT_SIZES['小四']
                    run_marker.font.superscript = True
                    run_marker.bold = False
                set_paragraph_spacing(cell.paragraphs[0])
            else:
                run = cell.paragraphs[0].add_run(str(cell_text))
                set_font(run, '宋体')
                run.font.size = FONT_SIZES['小四']
                run.bold = False
                set_paragraph_spacing(cell.paragraphs[0])

    # 表格注释（如果有）
    if note:
        note = replace_quote_tokens(note)
        note_para = doc.add_paragraph()
        note_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(note_para, FONT_SIZES['五号'])
        run = note_para.add_run(note)
        set_font(run, '宋体')
        run.font.size = FONT_SIZES['五号']
        run.bold = False

    # 脚注内容（如果有）
    if footnotes:
        # 脚注分隔：空半行
        fn_sep = doc.add_paragraph()
        set_paragraph_spacing(fn_sep, Pt(4))
        for marker, url in footnotes:
            fn_para = doc.add_paragraph()
            fn_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            set_paragraph_spacing(fn_para, FONT_SIZES['小五'])
            # 脚注标记
            run_marker = fn_para.add_run(marker + ' ')
            set_font(run_marker, '宋体')
            run_marker.font.size = FONT_SIZES['小五']
            run_marker.font.superscript = True
            # 超链接
            add_hyperlink_run(fn_para, url, url, '小五')
        # 空行分隔
        spacer = doc.add_paragraph()
        set_paragraph_spacing(spacer, Pt(2))

    return table


def set_default_style(doc):
    """设置文档默认样式：小四字体，行距固定 18 磅。"""
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = FONT_SIZES['小四']
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    pf = style.paragraph_format
    pf.line_spacing = Pt(18)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)


# ============================================================
# 封面处理
# ============================================================

# 中文数字映射
CN_DIGITS = {0: '○', 1: '一', 2: '二', 3: '三', 4: '四',
             5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}
CN_MONTHS = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六',
             7: '七', 8: '八', 9: '九', 10: '十', 11: '十一', 12: '十二'}


def _year_to_chinese(year: int) -> str:
    """将年份转为中文数字格式，如 2026 → 二○二六。"""
    return ''.join(CN_DIGITS.get(int(d), str(d)) for d in str(year))


def _replace_cover_placeholders(cover_doc, report_title: str):
    """替换封面模板中的占位符。

    【标题】 → 报告标题
    【二○二六】年 【三】月 → 当前日期中文格式
    """
    import re
    now = datetime.now()
    year_cn = _year_to_chinese(now.year)
    month_cn = CN_MONTHS.get(now.month, str(now.month))
    date_text = f'【{year_cn}】年 【{month_cn}】月'

    for p in cover_doc.paragraphs:
        full_text = p.text
        if not full_text.strip():
            continue

        # 替换标题占位符
        if '【标题】' in full_text:
            # 获取原始格式（取第一个 run 的格式）
            orig_font_size = None
            orig_bold = None
            orig_font_name = None
            if p.runs:
                orig_font_size = p.runs[0].font.size
                orig_bold = p.runs[0].bold
                orig_font_name = p.runs[0].font.name

            # 清除所有 run 文本，用第一个 run 承载新标题
            for run in p.runs:
                run.text = ''
            if p.runs:
                p.runs[0].text = report_title
            else:
                new_run = p.add_run(report_title)
                if orig_font_size:
                    new_run.font.size = orig_font_size
                if orig_bold is not None:
                    new_run.bold = orig_bold
                if orig_font_name:
                    new_run.font.name = orig_font_name
            continue

        # 替换日期占位符（含【…年…月】模式的段落）
        if re.search(r'【[^】]*\d[^】]*】\s*年', full_text) or \
           re.search(r'【[^】]*】\s*月', full_text):
            # 获取原始格式
            orig_font_size = None
            if p.runs:
                orig_font_size = p.runs[0].font.size

            # 清除所有 run，用第一个 run 承载新日期
            for run in p.runs:
                run.text = ''
            if p.runs:
                p.runs[0].text = date_text
                if orig_font_size:
                    p.runs[0].font.size = orig_font_size
            else:
                new_run = p.add_run(date_text)

    # 清理封面段落内嵌的分节符（sectPr in pPr），
    # 否则会和后续 add_section() 冲突，导致封面后出现空白页
    _remove_embedded_section_breaks(cover_doc)


def _remove_embedded_section_breaks(doc):
    """移除文档所有段落内嵌的分节符（w:sectPr inside w:pPr）。

    Word 封面模板常在末段插入分节符来控制封面页的版式。
    如果保留这些分节符，python-docx 在封面后 add_section() 时
    会产生一个空白过渡页。
    """
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for p in doc.paragraphs:
        pPr = p._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if pPr is not None:
            sectPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
            if sectPr is not None:
                pPr.remove(sectPr)


# ============================================================
# 主流程
# ============================================================

def generate_docx(json_path, output_path=None, cover_path=None):
    """从 JSON 文件生成 .docx 文档。

    参数：
        json_path: 报告 JSON 文件路径
        output_path: 输出 .docx 路径（可选）
        cover_path: 封面模板 .docx 路径（可选，将作为第一页）
    """
    # 读取 JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        content = json.load(f)

    # 违禁字符扫描
    violations = scan_json_content(content)
    if violations:
        print("=" * 60)
        print(" 违禁字符扫描未通过，禁止生成文档！")
        print("=" * 60)
        print(f"共发现 {len(violations)} 类违禁项：\n")
        for i, v in enumerate(violations, 1):
            print(f"{i}. {v['rule']}")
            for m in v['matches']:
                # 截断过长匹配
                display = m[:80] + ('...' if len(m) > 80 else '')
                print(f"   → {repr(display)}")
            print()
        print("请修改以上内容后重试。")
        return False

    print(" 违禁字符扫描通过，开始生成文档...")

    # 确定输出路径
    if output_path is None:
        title = content.get('title', '法律评估意见')
        date_str = datetime.now().strftime('%Y%m%d')
        output_path = f"{title}_{date_str}.docx"

    # 获取报告标题
    report_title = content.get('title', '法律评估意见书')

    # 创建文档（可选封面）
    has_cover = cover_path and os.path.exists(cover_path)
    if has_cover:
        doc = Document(cover_path)
        _replace_cover_placeholders(doc, report_title)
        # 封面后加分页符（不是分节符，不会在 Word 中显化分节标记）
        doc.add_page_break()
    else:
        doc = Document()

    set_default_style(doc)

    # 设置页面边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    # 生成标题（有封面时跳过，封面已展示标题）
    if not has_cover:
        add_title(doc, report_title)

    # 生成正文各节
    sections = content.get('sections', [])
    for sec in sections:
        sec_type = sec.get('type', 'body')
        text = replace_quote_tokens(sec.get('content', ''))

        if sec_type == 'h1':
            add_heading1(doc, text)
        elif sec_type == 'h2':
            add_heading2(doc, text)
        elif sec_type == 'h3':
            add_heading3(doc, text)
        elif sec_type == 'h4':
            add_heading4(doc, text)
        elif sec_type == 'h5':
            add_heading5(doc, text)
        elif sec_type == 'source_item':
            add_source_item(doc, text)
        elif sec_type == 'body':
            # 支持多段落（用 \n\n 分隔）
            paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
            if paragraphs:
                for p in paragraphs:
                    if '\n' in p:
                        add_body_with_line_breaks(doc, p)
                    else:
                        add_body(doc, p)
            else:
                if '\n' in text:
                    add_body_with_line_breaks(doc, text)
                else:
                    add_body(doc, text)
        elif sec_type == 'table':
            table_data = sec.get('content', {})
            fn_col = table_data.get('footnote_col')
            hl_col = table_data.get('hyperlink_col')
            fn_on_col = table_data.get('footnote_on_col')
            fn_urls = table_data.get('footnote_urls')
            add_table(
                doc,
                headers=table_data.get('headers', []),
                rows=table_data.get('rows', []),
                caption=table_data.get('caption'),
                note=table_data.get('note'),
                footnote_col=fn_col,
                hyperlink_col=hl_col,
                footnote_on_col=fn_on_col,
                footnote_urls=fn_urls
            )

    # 保存文档
    doc.save(output_path)
    abs_path = os.path.abspath(output_path)
    print(f" 文档已生成：{abs_path}")
    return True


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python3 generate_docx.py <input.json> [output.docx] [cover.docx]")
        print("示例：python3 generate_docx.py report.json 法律评估意见书.docx 封面.docx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    cover_file = sys.argv[3] if len(sys.argv) > 3 else None

    if not os.path.exists(input_file):
        print(f"错误：找不到输入文件 {input_file}")
        sys.exit(1)

    success = generate_docx(input_file, output_file, cover_file)
    sys.exit(0 if success else 1)
