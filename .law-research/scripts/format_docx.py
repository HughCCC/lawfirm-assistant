#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
格式调整校对引擎：读取任意 .docx，统一格式输出为标准法律文书格式。

用法：python3 format_docx.py <input.docx> [output.docx]
"""

import copy
import re
import sys
import os
from io import BytesIO

from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from docx.opc.part import Part
from docx.opc.packuri import PackURI


# ─── 猴子补丁：扩展 JPEG 识别 ──────────────────────────────────────
# python-docx 1.2.0 的 _ImageHeaderFactory 仅在 offset 6 处检测 "JFIF"
# 或 "Exif" 字符串来识别 JPEG。Adobe XMP / 部分 JPEG 变体在该位置是
# "http"（XMP 命名空间 URL），导致 UnrecognizedImageError。
# 补丁增加对 JPEG SOI 标记 (ff d8) 的通用检测，使所有 JPEG 变体可被识别。

def _patch_jpeg_recognition():
    """Monkey-patch python-docx's _ImageHeaderFactory to support all JPEG variants."""
    try:
        from docx.image.image import _ImageHeaderFactory as _orig_factory
        from docx.image.jpeg import _JfifMarkers, Jpeg
        from docx.image.exceptions import UnrecognizedImageError

        def _patched_factory(stream):
            stream.seek(0)
            header = stream.read(32)
            stream.seek(0)

            # Try original factory first (handles most standard formats)
            try:
                return _orig_factory(stream)
            except UnrecognizedImageError:
                pass

            # JPEG fallback: check SOI marker ff d8
            if header[0:2] == b'\xff\xd8':
                try:
                    markers = _JfifMarkers.from_stream(stream)
                    px_width = markers.sof.px_width
                    px_height = markers.sof.px_height
                    # Try Exif APP1 first, then JFIF APP0, default 72 DPI
                    try:
                        horz_dpi = markers.app1.horz_dpi
                        vert_dpi = markers.app1.vert_dpi
                    except Exception:
                        try:
                            horz_dpi = markers.app0.horz_dpi
                            vert_dpi = markers.app0.vert_dpi
                        except Exception:
                            horz_dpi, vert_dpi = 72, 72
                    return Jpeg(px_width, px_height, horz_dpi, vert_dpi)
                except Exception:
                    pass

            raise UnrecognizedImageError

        import docx.image.image
        docx.image.image._ImageHeaderFactory = _patched_factory
        return True
    except Exception:
        return False

_patch_jpeg_recognition()

# 复用 generate_docx 的格式化函数
from generate_docx import (
    set_default_style,
    set_paragraph_spacing,
    set_font,
    add_title,
    add_heading1, add_heading2, add_heading3, add_heading4, add_heading5,
    add_body, add_body_with_line_breaks, add_source_item,
    add_table,
    replace_quote_tokens,
    FONT_SIZES,
)


# ─── 图片保留 ────────────────────────────────────────────────────

def _has_image_in_para(para_element):
    """检查 OOXML 段落元素是否包含图片（drawing 或 pict）。"""
    for r_elem in para_element.findall(qn('w:r')):
        if r_elem.find(qn('w:drawing')) is not None:
            return True
        if r_elem.find(qn('w:pict')) is not None:
            return True
        # 兼容性图片（mc:AlternateContent → mc:Fallback → w:pict）
        ac = r_elem.find(
            '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent')
        if ac is not None:
            return True
    return False


def _copy_image_para(src_doc, doc, src_para):
    """复制含图片的段落：文本 run 正常复制，图片 run 提取 blob 后用 add_picture 重新插入。

    采用重建方式而非深层拷贝 XML，避免 rId 映射导致的 OPC package 兼容性问题。
    图片宽度约束到页面可用宽度（A4 21cm - 左右 3.18cm = 14.64cm），防止溢出。
    """
    from docx.shared import Inches, Emu

    # 页面可用宽度：A4(21cm) - 左(3.18cm) - 右(3.18cm) ≈ 14.6cm ≈ 5.75 inches
    MAX_WIDTH_INCHES = 5.7

    new_para = doc.add_paragraph()
    # 图片段落使用单倍行距 + 零段间距，否则固定 18pt 行距会裁切图片
    new_para.paragraph_format.line_spacing = 1.0
    new_para.paragraph_format.space_before = Pt(0)
    new_para.paragraph_format.space_after = Pt(0)
    # 保留段落对齐方式
    if src_para.alignment is not None:
        new_para.alignment = src_para.alignment

    src_part = src_doc.part

    for run in src_para.runs:
        has_image = False
        # 检查 w:drawing
        drawing = run._element.find(qn('w:drawing'))
        # 检查 w:pict
        pict = run._element.find(qn('w:pict'))
        # 检查 mc:AlternateContent
        ac = run._element.find(
            '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent')

        if drawing is not None:
            has_image = True
            # 提取图片 rId → blob → 重新插入
            blip_elems = drawing.findall('.//' + qn('a:blip'))
            for blip in blip_elems:
                rId = blip.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rId and rId in src_part.rels:
                    rel = src_part.rels[rId]
                    if 'image' in rel.reltype:
                        blob = rel.target_part.blob
                        try:
                            ext = drawing.find('.//' + qn('wp:extent'))
                            if ext is not None:
                                cx = int(ext.get('cx', 0))
                                cy = int(ext.get('cy', 0))
                                # 约束宽度到页面可用范围，避免图像超出右边界
                                width_inches = Emu(cx).inches if cx else 0
                                height_inches = Emu(cy).inches if cy else 0
                                if width_inches > MAX_WIDTH_INCHES:
                                    width_inches = MAX_WIDTH_INCHES
                                if width_inches > 0:
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(blob), width=Inches(width_inches))
                                else:
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(blob))
                            else:
                                run_pic = new_para.add_run()
                                run_pic.add_picture(BytesIO(blob))
                        except Exception:
                            try:
                                run_pic = new_para.add_run()
                                run_pic.add_picture(BytesIO(blob))
                            except Exception:
                                pass  # 最终兜底：图片格式不被识别，静默跳过
        elif ac is not None:
            has_image = True
            # AlternateContent 中的图片：尝试提取 rId + 约束宽度
            for elem in ac.iter():
                blip_elems = elem.findall(qn('a:blip'))
                for blip in blip_elems:
                    rId = blip.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rId and rId in src_part.rels:
                        rel = src_part.rels[rId]
                        if 'image' in rel.reltype:
                            try:
                                run_pic = new_para.add_run()
                                run_pic.add_picture(BytesIO(rel.target_part.blob),
                                                    width=Inches(MAX_WIDTH_INCHES))
                            except Exception:
                                pass
        elif pict is not None:
            has_image = True
            # w:pict (旧格式) — 提取图片数据 + 约束宽度
            for imagedata in pict.findall('.//' + qn('v:imagedata')):
                rId = imagedata.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rId and rId in src_part.rels:
                    rel = src_part.rels[rId]
                    if 'image' in rel.reltype:
                        try:
                            run_pic = new_para.add_run()
                            run_pic.add_picture(BytesIO(rel.target_part.blob),
                                                width=Inches(MAX_WIDTH_INCHES))
                        except Exception:
                            pass

        if not has_image and run.text and run.text.strip():
            # 普通文本 run
            new_run = new_para.add_run(run.text)
            set_font(new_run, '宋体')
            if run.bold:
                new_run.bold = True

    return new_para


# ─── 脚注/尾注处理 ───────────────────────────────────────────────


def _para_has_footnote_ref(para_element):
    """检查 OOXML 段落元素是否包含脚注或尾注引用。"""
    for r_elem in para_element.findall(qn('w:r')):
        if r_elem.find(qn('w:footnoteReference')) is not None:
            return True
        if r_elem.find(qn('w:endnoteReference')) is not None:
            return True
    return False


def _format_notes_blob(blob, ref_tag, type_attr):
    """解析脚注/尾注 XML blob，标准化格式。

    1. 引用标记（w:footnoteRef/w:endnoteRef）强制使用内置样式 +
       显式设置 w:vertAlign="superscript" 和 小五字号作为直接格式兜底
    2. 内容文本格式化为宋体小五 (9pt)
    3. 保留已有标准字符样式（FootnoteReference/EndnoteReference）的 run 不变

    采用重建 rPr 策略：对所有 run 先清空 rPr 子元素再按序重建，
    杜绝深拷贝残留属性（bold、color、多余 rFonts attrs 等）泄漏。
    """
    from lxml import etree

    if not blob or not blob.strip():
        print("警告：脚注/尾注 XML blob 为空，跳过格式化")
        return blob  # 原样返回，不处理

    try:
        root = etree.fromstring(blob)
    except etree.XMLSyntaxError as e:
        print(f"错误：脚注/尾注 XML 解析失败 — {e}")
        return blob  # 原样返回，让输出文档至少保留原始脚注

    # 根据引用类型确定标准样式名
    is_footnote = ('footnoteRef' in ref_tag)
    standard_style = 'FootnoteReference' if is_footnote else 'EndnoteReference'

    for note_elem in root:
        # 跳过分隔符和连续分隔符（无正文内容）
        note_type = note_elem.get(qn(type_attr))
        if note_type in ('separator', 'continuationSeparator'):
            continue

        # 遍历该脚注/尾注中的所有 w:p → w:r 元素
        for p_elem in note_elem:
            if p_elem.tag != qn('w:p'):
                continue
            # 使用 .iter(qn('w:r')) 递归查找所有 w:r 元素，
            # 包括嵌套在 w:hyperlink 等元素内的 run
            for r_elem in p_elem.iter(qn('w:r')):

                # 引用标记 run（w:footnoteRef / w:endnoteRef）：强制使用内置标准样式
                if r_elem.find(qn(ref_tag)) is not None:
                    # 找到或创建 rPr，然后清空所有子元素
                    rPr = r_elem.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = etree.Element(qn('w:rPr'))
                        r_elem.insert(0, rPr)
                    else:
                        for child in list(rPr):
                            rPr.remove(child)

                    # 1. rStyle — 标准内置样式
                    rStyle = etree.Element(qn('w:rStyle'))
                    rPr.append(rStyle)
                    rStyle.set(qn('w:val'), standard_style)

                    # 2. vertAlign — 上标
                    vertAlign = etree.Element(qn('w:vertAlign'))
                    rPr.append(vertAlign)
                    vertAlign.set(qn('w:val'), 'superscript')

                    # 3. sz / szCs — 小五 (9pt = 18 half-points)
                    sz = etree.Element(qn('w:sz'))
                    rPr.append(sz)
                    sz.set(qn('w:val'), '18')

                    szCs = etree.Element(qn('w:szCs'))
                    rPr.append(szCs)
                    szCs.set(qn('w:val'), '18')
                    continue

                # 跳过含标准内置字符样式的 run（如 FootnoteReference/EndnoteReference）
                rPr = r_elem.find(qn('w:rPr'))
                if rPr is not None:
                    rStyle = rPr.find(qn('w:rStyle'))
                    if rStyle is not None:
                        style_val = rStyle.get(qn('w:val'))
                        # 保留标准样式，但覆盖自定义样式
                        if style_val in ('FootnoteReference', 'EndnoteReference'):
                            continue

                # 普通文本 run：清除旧格式，重建为宋体小五
                if rPr is None:
                    rPr = etree.Element(qn('w:rPr'))
                    r_elem.insert(0, rPr)
                else:
                    # 清空所有旧子元素，从零重建
                    for child in list(rPr):
                        rPr.remove(child)

                # 1. rFonts — 宋体 + Times New Roman（清除旧 rFonts 可能残留的 hAnsi/cs 等属性）
                rFonts = etree.Element(qn('w:rFonts'))
                rPr.append(rFonts)
                rFonts.set(qn('w:eastAsia'), '宋体')
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                # 显式清除西文字体其他槽位，防止残留
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rFonts.set(qn('w:cs'), 'Times New Roman')

                # 2. sz / szCs — 小五 (9pt = 18 half-points)
                sz = etree.Element(qn('w:sz'))
                rPr.append(sz)
                sz.set(qn('w:val'), '18')

                szCs = etree.Element(qn('w:szCs'))
                rPr.append(szCs)
                szCs.set(qn('w:val'), '18')

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _copy_footnotes_part(src_doc, dst_doc):
    """将源文档的脚注部件（word/footnotes.xml）复制到目标文档，
    并将脚注文本格式化为宋体小五 (9pt)。"""
    try:
        src_part = src_doc.part.part_related_by(RT.FOOTNOTES)
    except KeyError:
        return False
    # 解析并格式化脚注内容
    formatted_blob = _format_notes_blob(src_part.blob, 'w:footnoteRef', 'w:type')
    package = dst_doc.part.package
    partname = PackURI('/word/footnotes.xml')
    new_part = Part(partname, CT.WML_FOOTNOTES, formatted_blob, package)
    dst_doc.part.relate_to(new_part, RT.FOOTNOTES)
    return True


def _copy_endnotes_part(src_doc, dst_doc):
    """将源文档的尾注部件（word/endnotes.xml）复制到目标文档，
    并将尾注文本格式化为宋体小五 (9pt)。"""
    try:
        src_part = src_doc.part.part_related_by(RT.ENDNOTES)
    except KeyError:
        return False
    # 解析并格式化尾注内容
    formatted_blob = _format_notes_blob(src_part.blob, 'w:endnoteRef', 'w:type')
    package = dst_doc.part.package
    partname = PackURI('/word/endnotes.xml')
    new_part = Part(partname, CT.WML_ENDNOTES, formatted_blob, package)
    dst_doc.part.relate_to(new_part, RT.ENDNOTES)
    return True


def _apply_para_formatting(para, ptype):
    """根据段落类型应用标准段落格式（对齐、行距、缩进）。"""
    from docx.shared import Pt as _Pt

    pf = para.paragraph_format
    pf.line_spacing = _Pt(18)
    pf.space_before = _Pt(6)
    pf.space_after = _Pt(6)

    if ptype == 'title':
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif ptype in ('h1', 'h2', 'source_item'):
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif ptype in ('h3', 'h4', 'h5', 'body'):
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.first_line_indent = _Pt(24)


def _apply_rPr_style(r_elem, small_five=False):
    """对 w:r 元素应用标准字体格式。

    保留已有的 w:rStyle（如 FootnoteReference/EndnoteReference），
    保留含脚注引用/尾注引用（w:footnoteReference/w:endnoteReference）的 run 不变。
    仅覆盖 w:rFonts 和 w:sz，确保脚注/尾注引用保持上标等特殊样式。

    Args:
        r_elem: w:r OOXML 元素
        small_five: True 则使用小五 (9pt)，默认小四 (12pt)
    """
    # 含脚注/尾注引用的 run — 完全不修改，保持源文档的 superscript 等格式
    if (r_elem.find(qn('w:footnoteReference')) is not None or
            r_elem.find(qn('w:endnoteReference')) is not None):
        return

    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = r_elem.makeelement(qn('w:rPr'), {})
        r_elem.insert(0, rPr)

    # 保留已有的 rStyle（FootnoteReference 等），不覆盖字体和大小
    rStyle = rPr.find(qn('w:rStyle'))
    if rStyle is not None:
        return

    # 设置中英文字体
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), '宋体')
    rFonts.set(qn('w:ascii'), 'Times New Roman')

    # 设置字号：小四12pt (= 24 half-points) 或 小五9pt (= 18 half-points)
    sz_val = '18' if small_five else '24'
    sz = rPr.find(qn('w:sz'))
    if sz is None:
        sz = rPr.makeelement(qn('w:sz'), {})
        rPr.append(sz)
    sz.set(qn('w:val'), sz_val)

    szCs = rPr.find(qn('w:szCs'))
    if szCs is None:
        szCs = rPr.makeelement(qn('w:szCs'), {})
        rPr.append(szCs)
    szCs.set(qn('w:val'), sz_val)


def _normalize_fn_ref_run(r_elem):
    """将脚注/尾注引用的 w:r 元素标准化为内置样式。

    源文档可能使用自定义样式（如 af0），新文档中没有该样式会导致上标丢失。
    此函数强制使用 Word 内置的 FootnoteReference / EndnoteReference 样式，
    并显式设置 w:vertAlign="superscript" + 字号作为直接格式兜底，
    因为 python-docx 默认模板中的这些样式可能不含上标/字号定义。

    采用重建 rPr 策略：先删除所有旧 rPr 子元素，再按正确顺序
    插入所需元素，杜绝深拷贝残留属性（bold、color、font 等）泄漏。

    """
    # 确定引用类型
    is_footnote = r_elem.find(qn('w:footnoteReference')) is not None
    style_val = 'FootnoteReference' if is_footnote else 'EndnoteReference'

    # 找到或创建 rPr，然后清空所有子元素（杜绝深拷贝残留）
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = r_elem.makeelement(qn('w:rPr'), {})
        r_elem.insert(0, rPr)
    else:
        # 清空所有旧子元素，从零重建
        for child in list(rPr):
            rPr.remove(child)

    # 1. w:rStyle — 标准内置样式（必须在首位）
    rStyle = rPr.makeelement(qn('w:rStyle'), {})
    rPr.append(rStyle)
    rStyle.set(qn('w:val'), style_val)

    # 2. w:vertAlign — 上标（直接格式，即使样式缺失也能正确显示）
    vertAlign = rPr.makeelement(qn('w:vertAlign'), {})
    rPr.append(vertAlign)
    vertAlign.set(qn('w:val'), 'superscript')

    # 3. w:sz / w:szCs — 字号（正文引用用小四 12pt=24，脚注区域用小五 9pt=18）
    font_size = '18'  # 默认小五，脚注区域引用
    sz = rPr.makeelement(qn('w:sz'), {})
    rPr.append(sz)
    sz.set(qn('w:val'), font_size)

    szCs = rPr.makeelement(qn('w:szCs'), {})
    rPr.append(szCs)
    szCs.set(qn('w:val'), font_size)


def _copy_footnote_para(doc, src_para, ptype=None):
    """深拷贝含脚注/尾注引用的段落，应用标准格式。

    对于含 w:footnoteReference 或 w:endnoteReference 的 run，深拷贝整个
    w:r 元素以保留引用结构，并强制使用内置 FootnoteReference /
    EndnoteReference 样式确保上标正常显示。
    对于纯文本 run，创建新的标准化 w:r 元素。
    严格保持源段落的 run 顺序。
    """
    new_para = doc.add_paragraph()
    _apply_para_formatting(new_para, ptype)

    p_elem = new_para._element

    # 移除 add_paragraph() 自动创建的空 run
    for child in list(p_elem):
        if child.tag == qn('w:r'):
            has_fn = child.find(qn('w:footnoteReference')) is not None
            has_en = child.find(qn('w:endnoteReference')) is not None
            has_drawing = child.find(qn('w:drawing')) is not None
            if not (has_fn or has_en or has_drawing):
                text_elems = child.findall(qn('w:t'))
                has_text = any(t.text and t.text.strip() for t in text_elems)
                if not has_text:
                    p_elem.remove(child)

    # 按源段落 run 顺序逐一复制
    for run in src_para.runs:
        r_elem = run._element
        has_fn = r_elem.find(qn('w:footnoteReference')) is not None
        has_en = r_elem.find(qn('w:endnoteReference')) is not None

        if has_fn or has_en:
            # 脚注/尾注引用 run：深拷贝后强制使用内置样式确保上标
            new_r = copy.deepcopy(r_elem)
            _normalize_fn_ref_run(new_r)
            p_elem.append(new_r)
        else:
            # 纯文本 run：创建新的标准化 w:r
            text = run.text
            if text:
                new_r = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:eastAsia'), '宋体')
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rPr.append(rFonts)
                sz_elem = OxmlElement('w:sz')
                sz_elem.set(qn('w:val'), '24')
                rPr.append(sz_elem)
                if run.bold:
                    b_elem = OxmlElement('w:b')
                    rPr.append(b_elem)
                new_r.append(rPr)
                t_elem = OxmlElement('w:t')
                t_elem.text = text
                t_elem.set(qn('xml:space'), 'preserve')
                new_r.append(t_elem)
                p_elem.append(new_r)

    return new_para


def _cell_has_footnote_ref(cell):
    """检查表格单元格中任一段落是否含脚注/尾注引用。"""
    for para in cell.paragraphs:
        if _para_has_footnote_ref(para._element):
            return True
    return False


def _cell_has_image(cell):
    """检查表格单元格中任一段落是否包含图片。"""
    for para in cell.paragraphs:
        if _has_image_in_para(para._element):
            return True
    return False


def _copy_cell_para(dst_cell, src_para, src_doc, is_header=False):
    """在表格单元格中复制单个段落，保留文本/图片/脚注引用。

    根据段落内容类型选择复制策略：
    - 含图片：使用 _copy_image_para 重建（但将结果移入单元格）
    - 含脚注/尾注：深拷贝引用结构 + 标准化格式
    - 纯文本：创建新的标准化文本 run
    """
    from docx.oxml import OxmlElement as OE

    has_image = _has_image_in_para(src_para._element)
    has_fn = _para_has_footnote_ref(src_para._element)
    text = src_para.text.strip()

    # 如果段落既无文本也无图片，跳过
    if not text and not has_image and not has_fn:
        return

    if has_image:
        # 含图片的段落：在单元格内用 add_picture 重建
        from docx.shared import Inches, Emu
        MAX_WIDTH_INCHES = 4.5  # 表格内图片宽度约束更严格

        new_para = dst_cell.add_paragraph()
        new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        new_para.paragraph_format.line_spacing = 1.0
        new_para.paragraph_format.space_before = Pt(0)
        new_para.paragraph_format.space_after = Pt(0)

        src_part = src_doc.part
        for run in src_para.runs:
            run_elem = run._element
            drawing = run_elem.find(qn('w:drawing'))
            pict = run_elem.find(qn('w:pict'))
            ac = run_elem.find(
                '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent')

            if drawing is not None:
                blip_elems = drawing.findall('.//' + qn('a:blip'))
                for blip in blip_elems:
                    rId = blip.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rId and rId in src_part.rels:
                        rel = src_part.rels[rId]
                        if 'image' in rel.reltype:
                            blob = rel.target_part.blob
                            try:
                                ext = drawing.find('.//' + qn('wp:extent'))
                                if ext is not None:
                                    cx = int(ext.get('cx', 0))
                                    width_inches = Emu(cx).inches if cx else 0
                                    if width_inches > MAX_WIDTH_INCHES or width_inches <= 0:
                                        width_inches = MAX_WIDTH_INCHES
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(blob), width=Inches(width_inches))
                                else:
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(blob))
                            except Exception:
                                try:
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(blob))
                                except Exception:
                                    pass  # 最终兜底：图片格式不被识别，静默跳过
            elif ac is not None:
                for elem in ac.iter():
                    blip_elems = elem.findall(qn('a:blip'))
                    for blip in blip_elems:
                        rId = blip.get(
                            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if rId and rId in src_part.rels:
                            rel = src_part.rels[rId]
                            if 'image' in rel.reltype:
                                try:
                                    run_pic = new_para.add_run()
                                    run_pic.add_picture(BytesIO(rel.target_part.blob),
                                                        width=Inches(MAX_WIDTH_INCHES))
                                except Exception:
                                    pass
            elif pict is not None:
                for imagedata in pict.findall('.//' + qn('v:imagedata')):
                    rId = imagedata.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rId and rId in src_part.rels:
                        rel = src_part.rels[rId]
                        if 'image' in rel.reltype:
                            try:
                                run_pic = new_para.add_run()
                                run_pic.add_picture(BytesIO(rel.target_part.blob),
                                                    width=Inches(MAX_WIDTH_INCHES))
                            except Exception:
                                pass
            elif run.text and run.text.strip():
                new_run = new_para.add_run(run.text)
                set_font(new_run, '宋体')
    elif has_fn:
        # 含脚注/尾注引用：使用深拷贝 + 标准化
        _copy_footnote_para_to_cell(dst_cell, src_para, is_header=is_header)
    else:
        # 纯文本段落
        new_para = dst_cell.add_paragraph()
        run = new_para.add_run(text)
        set_font(run, '宋体')
        run.font.size = FONT_SIZES['小四']
        if is_header:
            run.bold = True
        new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(new_para, FONT_SIZES['小四'])


def _copy_table_with_content(doc, src_doc, src_table):
    """重建表格，保留所有内容（文本、图片、脚注引用）并应用标准格式。"""
    from docx.enum.table import WD_TABLE_ALIGNMENT

    num_rows = len(src_table.rows)
    num_cols = len(src_table.columns)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    for r_idx, src_row in enumerate(src_table.rows):
        for c_idx, src_cell in enumerate(src_row.cells):
            try:
                dst_cell = table.rows[r_idx].cells[c_idx]

                # 移除目标单元格中的默认空段落
                for p in list(dst_cell.paragraphs):
                    p_elem = p._element
                    p_elem.getparent().remove(p_elem)

                # 复制源单元格内容（逐段落）
                for src_para in src_cell.paragraphs:
                    _copy_cell_para(dst_cell, src_para, src_doc, is_header=(r_idx == 0))
            except Exception as e:
                # 单个单元格失败不影响同表其他单元格
                print(f"警告：表格单元格 [{r_idx},{c_idx}] 处理失败 — {e}，已跳过")

    return table


# ─── 图片注解检测 ─────────────────────────────────────────────────


def _is_image_caption(text):
    """判断文本是否为图片注解（含括号的说明性文字）。

    检测中文全角括号（）或英文半角括号 ()，视为图片下方注解文字。
    """
    if not text:
        return False
    # 中文全角括号
    if '（' in text and '）' in text:
        return True
    # 英文半角括号（排除纯 URL 链接）
    if '(' in text and ')' in text:
        # 排除以 http/https 开头的纯链接
        if not re.match(r'^https?://', text.strip()):
            return True
    return False


def _add_image_caption(doc, text):
    """添加图片注解：宋体五号居中。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(para, FONT_SIZES['五号'])
    run = para.add_run(replace_quote_tokens(text))
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['五号']
    run.bold = False
    return para


# ─── 表格注释检测 ─────────────────────────────────────────────────

# 表格注释常见前缀模式
_TABLE_NOTE_PATTERNS = [
    re.compile(r'^(来源|注|表注|数据来源|资料来源|以上数据|备注|说明)[：:]'),
    re.compile(r'^\*'),  # 星号注释（如 *p < 0.05）
]


def _is_table_note(text):
    """判断文本是否为表格下方注释（以常见注释前缀开头）。"""
    if not text:
        return False
    for pat in _TABLE_NOTE_PATTERNS:
        if pat.match(text.strip()):
            return True
    return False


def _add_table_note(doc, text):
    """添加表格注释：宋体五号居中。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(para, FONT_SIZES['五号'])
    run = para.add_run(replace_quote_tokens(text))
    set_font(run, '宋体')
    run.font.size = FONT_SIZES['五号']
    run.bold = False
    return para


def _copy_footnote_para_to_cell(cell, src_para, is_header=False):
    """在表格单元格中创建含脚注/尾注引用的段落，保持引用结构。"""
    new_para = cell.add_paragraph()
    new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(new_para, FONT_SIZES['小四'])

    p_elem = new_para._element

    # 移除 add_paragraph() 自动创建的空 run
    for child in list(p_elem):
        if child.tag == qn('w:r'):
            has_special = (child.find(qn('w:footnoteReference')) is not None or
                          child.find(qn('w:endnoteReference')) is not None or
                          child.find(qn('w:drawing')) is not None)
            if not has_special:
                text_elems = child.findall(qn('w:t'))
                has_text = any(t.text and t.text.strip() for t in text_elems)
                if not has_text:
                    p_elem.remove(child)

    # 按源段落 run 顺序逐一复制
    for run in src_para.runs:
        r_elem = run._element
        has_fn = r_elem.find(qn('w:footnoteReference')) is not None
        has_en = r_elem.find(qn('w:endnoteReference')) is not None

        if has_fn or has_en:
            # 脚注/尾注引用 run：深拷贝后强制使用内置样式确保上标
            new_r = copy.deepcopy(r_elem)
            _normalize_fn_ref_run(new_r)
            p_elem.append(new_r)
        else:
            text = run.text
            if text:
                new_r = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:eastAsia'), '宋体')
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rPr.append(rFonts)
                sz_elem = OxmlElement('w:sz')
                sz_elem.set(qn('w:val'), '24')
                rPr.append(sz_elem)
                if run.bold or is_header:
                    b_elem = OxmlElement('w:b')
                    rPr.append(b_elem)
                new_r.append(rPr)
                t_elem = OxmlElement('w:t')
                t_elem.text = text
                t_elem.set(qn('xml:space'), 'preserve')
                new_r.append(t_elem)
                p_elem.append(new_r)

    return new_para


def _copy_table_with_fn(doc, src_table):
    """重建含脚注/尾注引用的表格，逐单元格保留引用结构并应用标准格式。"""
    from docx.enum.table import WD_TABLE_ALIGNMENT

    num_rows = len(src_table.rows)
    num_cols = len(src_table.columns)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    for r_idx, src_row in enumerate(src_table.rows):
        for c_idx, src_cell in enumerate(src_row.cells):
            dst_cell = table.rows[r_idx].cells[c_idx]

            # 移除目标单元格中的默认空段落
            for p in list(dst_cell.paragraphs):
                p_elem = p._element
                p_elem.getparent().remove(p_elem)

            # 复制源单元格内容
            for src_para in src_cell.paragraphs:
                text = src_para.text.strip()
                has_fn = _para_has_footnote_ref(src_para._element)

                if not text and not has_fn:
                    continue

                if has_fn:
                    _copy_footnote_para_to_cell(dst_cell, src_para, is_header=(r_idx == 0))
                else:
                    new_para = dst_cell.add_paragraph()
                    run = new_para.add_run(text)
                    set_font(run, '宋体')
                    run.font.size = FONT_SIZES['小四']
                    if r_idx == 0:
                        run.bold = True
                    new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_paragraph_spacing(new_para, FONT_SIZES['小四'])

    return table


# ─── 段落类型检测 ─────────────────────────────────────────────────

# 中文序号 → h1
RE_H1 = re.compile(r'^[一二三四五六七八九十]、')

# 括号中文序号 → h2
RE_H2 = re.compile(r'^（[一二三四五六七八九十]+）')

# 阿拉伯数字 + 、→ h3
RE_H3 = re.compile(r'^\d+、')

# 括号阿拉伯数字 → h4
RE_H4 = re.compile(r'^（\d+）')

# 圈号数字 → h5
RE_H5 = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]')


def detect_paragraph_type(para, is_first=False):
    """检测段落的逻辑类型。

    优先级：内容序号模式 > 标题检测（首段居中/bold） > bold/indent 辅助判断 > 默认 body
    返回: (type_str, text_content)
    """
    text = para.text.strip()
    if not text:
        return None, text  # 空段落跳过

    # 序号模式匹配（优先级最高）
    if RE_H1.match(text):
        return 'h1', text
    if RE_H2.match(text):
        return 'h2', text
    if RE_H3.match(text):
        return 'h3', text
    if RE_H4.match(text):
        return 'h4', text
    if RE_H5.match(text):
        return 'h5', text

    # 标题检测：首个非空段落 + 无编号前缀 + (居中 或 bold)
    if is_first:
        alignment = para.alignment
        is_bold = _is_paragraph_bold(para)
        if alignment == WD_ALIGN_PARAGRAPH.CENTER or is_bold:
            return 'title', text

    # 辅助判断：bold + 无首行缩进 → 可能为标题
    is_bold = _is_paragraph_bold(para)
    has_indent = _has_first_line_indent(para)

    if is_bold and not has_indent:
        # 可能是 h1/h2/source_item
        return 'h2', text  # 默认按 h2 处理
    if is_bold and has_indent:
        return 'h3', text  # 默认按 h3 处理

    # 默认 body
    return 'body', text


def _is_paragraph_bold(para):
    """判断段落第一个 run 是否加粗。"""
    for run in para.runs:
        if run.text.strip():
            return run.bold is True
    return False


def _has_first_line_indent(para):
    """判断段落是否有首行缩进（约 24pt）。"""
    indent = para.paragraph_format.first_line_indent
    if indent is None:
        return False
    # 24pt = 342900 EMU, 允许一定容差
    return indent > 200000  # 约 5.5pt 以上视为有缩进


def _get_alignment(para):
    """获取段落对齐方式。"""
    return para.alignment


# ─── 表格提取 ────────────────────────────────────────────────────

def _extract_table_data(table):
    """从原始表格提取 headers + rows。"""
    headers = []
    rows = []

    for i, row in enumerate(table.rows):
        cells = [cell.text.strip() for cell in row.cells]
        if i == 0:
            headers = cells
        else:
            rows.append(cells)

    return headers, rows


# ─── 主流程 ─────────────────────────────────────────────────────

def format_docx(input_path, output_path=None):
    """读取 input.docx，格式化输出为标准法律文书格式。

    参数：
        input_path:  输入 .docx 文件路径
        output_path: 输出 .docx 文件路径（可选，默认加 _formatted 后缀）

    返回：
        (success: bool, output_path: str)
    """
    if not os.path.exists(input_path):
        print(f"错误：找不到输入文件 {input_path}")
        return False, input_path

    # 确定输出路径
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_formatted{ext}"

    print(f"正在读取：{input_path}")

    # 1. 读取原始文档
    try:
        src_doc = Document(input_path)
    except Exception as e:
        print(f"错误：无法读取文档 — {e}")
        return False, input_path

    # 2. 创建新文档
    doc = Document()
    set_default_style(doc)

    # 页面边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    # 2.5 复制脚注/尾注部件（如有）
    try:
        has_footnotes = _copy_footnotes_part(src_doc, doc)
    except Exception as e:
        print(f"错误：处理脚注时失败 — {e}")
        return False, f"脚注处理失败：{e}"
    try:
        has_endnotes = _copy_endnotes_part(src_doc, doc)
    except Exception as e:
        print(f"错误：处理尾注时失败 — {e}")
        return False, f"尾注处理失败：{e}"
    if has_footnotes:
        print("  已保留脚注内容")
    if has_endnotes:
        print("  已保留尾注内容")

    # 3. 按文档顺序遍历所有元素（段落和表格交错处理）
    body = src_doc.element.body
    para_count = 0
    table_count = 0
    skipped_empty = 0
    first_content_seen = False
    prev_element = None  # 'image' | 'table' | None（用于检测图片注解/表格注释）

    for child in body:
        try:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                # 段落元素
                if para_count < len(src_doc.paragraphs):
                    para = src_doc.paragraphs[para_count]
                else:
                    para_count += 1
                    continue
                para_count += 1

                # 检查是否为上方图片的注解文字（含括号）
                if prev_element == 'image':
                    ptype, text = detect_paragraph_type(para, is_first=(not first_content_seen))
                    if ptype is not None and _is_image_caption(text):
                        _add_image_caption(doc, text)
                        first_content_seen = True
                        prev_element = None
                        continue
                    prev_element = None

                # 检查是否为上方表格的注释文字
                if prev_element == 'table':
                    ptype, text = detect_paragraph_type(para, is_first=(not first_content_seen))
                    if ptype is not None and _is_table_note(text):
                        _add_table_note(doc, text)
                        first_content_seen = True
                        prev_element = None
                        continue
                    prev_element = None

                # 图片段落：保留原样（深层拷贝 + rId 迁移）
                if _has_image_in_para(child):
                    _copy_image_para(src_doc, doc, para)
                    first_content_seen = True
                    prev_element = 'image'
                    continue

                # 脚注/尾注段落：深层拷贝以保留引用结构
                if _para_has_footnote_ref(child):
                    ptype, text = detect_paragraph_type(para, is_first=(not first_content_seen))
                    if ptype is not None:
                        _copy_footnote_para(doc, para, ptype)
                        first_content_seen = True
                    else:
                        skipped_empty += 1
                    prev_element = None
                    continue

                # 普通段落：重置前驱元素类型
                prev_element = None

                ptype, text = detect_paragraph_type(para, is_first=(not first_content_seen))

                if ptype is None:
                    skipped_empty += 1
                    continue

                first_content_seen = True

                # 替换 LQ/RQ token（如有）
                text = replace_quote_tokens(text)

                # 按类型写入
                if ptype == 'title':
                    add_title(doc, text)
                elif ptype == 'h1':
                    add_heading1(doc, text)
                elif ptype == 'h2':
                    add_heading2(doc, text)
                elif ptype == 'h3':
                    add_heading3(doc, text)
                elif ptype == 'h4':
                    add_heading4(doc, text)
                elif ptype == 'h5':
                    add_heading5(doc, text)
                elif ptype == 'body':
                    # 段内换行检测
                    if '\n' in text:
                        add_body_with_line_breaks(doc, text)
                    else:
                        add_body(doc, text)

            elif tag == 'tbl':
                # 表格元素
                if table_count < len(src_doc.tables):
                    table = src_doc.tables[table_count]
                    table_count += 1

                    # 检查是否有单元格含脚注/尾注引用或图片
                    has_fn = any(
                        _cell_has_footnote_ref(cell)
                        for row in table.rows for cell in row.cells
                    )
                    has_img = any(
                        _cell_has_image(cell)
                        for row in table.rows for cell in row.cells
                    )

                    if has_fn or has_img:
                        # 含脚注或图片：使用内容保留复制，确保文本/图片/脚注均不丢失
                        _copy_table_with_content(doc, src_doc, table)
                    else:
                        # 纯文本表格：轻量路径，提取 headers/rows
                        headers, rows = _extract_table_data(table)
                        if headers and rows:
                            add_table(
                                doc,
                                headers=headers,
                                rows=rows,
                                caption=None,
                                note=None,
                            )
                    prev_element = 'table'
        except Exception as e:
            elem_desc = f"第{para_count}段" if tag == 'p' else f"第{table_count + 1}个表格"
            err_msg = str(e) if str(e) else f"未知错误（{type(e).__name__}）"
            print(f"警告：{elem_desc}处理失败 — {err_msg}，已跳过")
            prev_element = None
            continue

    # 4. 保存
    try:
        doc.save(output_path)
    except Exception as e:
        print(f"错误：保存文档失败 — {e}")
        return False, f"保存失败：{e}"
    abs_path = os.path.abspath(output_path)

    print(f"段落: {para_count} | 表格: {table_count} | 跳过空段: {skipped_empty}")
    print(f"文档已生成：{abs_path}")
    return True, abs_path


# ─── CLI ─────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python3 format_docx.py <input.docx> [output.docx]")
        print("示例：python3 format_docx.py 混乱格式.docx 标准格式.docx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    success, _ = format_docx(input_file, output_file)
    sys.exit(0 if success else 1)
