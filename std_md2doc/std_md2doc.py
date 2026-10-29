# -*- coding: utf-8 -*-
"""Markdown → DOCX 转换器（自建版式，不依赖任何外部模板文件）。

用法：
    python local3.py plan-a.md [plan-a.docx]

版式全部由脚本在空白文档上现场生成（不再读取《综合设计文档_V0.6.docx》）：
    * A4 页面、上下左右边距 2.5/2.0/2.5/2.5 cm、页脚距下端 1.5 cm；
    * 正文 Normal = 微软雅黑 10.5pt、行距 1.21、段后 10pt、颜色 2B3A4A、左对齐；
      （纸张/边距/字号/行距/对齐均可用下面五个命令行参数覆盖）
    * 标题样式：Heading 1（封面主标题 28pt，居中）/ Heading 2（章 15.5pt）/ Heading 3（节 12.5pt）/
      Heading 4（10.5pt，颜色 4F81BD），均粗体、与下段同页、除封面外左对齐；
    * 表格：1px C9D4DE 边框、表头 EAF1F8 底纹 + 1F4E79 粗体、单元格垂直居中、表头跨页重复；
    * 封面页 / 版本历史页 / 目录页 / 正文页四分节，仅正文分节显示居中页码并从 1 重新编号。

版式可覆盖（不改代码即可适配其它文档规范）：
    --page-size A4|A3|Letter   纸张尺寸，缺省 A4（21.0 × 29.7 cm）
    --margin 上,下,左,右       页边距，单位 cm，缺省 2.5,2.0,2.5,2.5；也可只给 1 个或 2 个值
    --body-pt 10.5             正文字号（磅），缺省 10.5
    --line-spacing 1.21        正文行距（倍数），缺省 1.21；Heading 4 的行距跟随此值
    --justify                  正文两端对齐（缺省左对齐）
    覆盖值写回分节与 Normal 样式后，text_width_emu、表格列宽、封面标题可排宽度
    与图片最大宽度都按新版心重算；不带这些参数时输出与旧版逐字节一致。

示例：
    python local3.py plan-a.md out.docx
    python local3.py plan-a.md out.docx --page-size A3 --margin 2.5,2.5,3,3
    python local3.py plan-a.md out.docx --body-pt 12 --line-spacing 1.5 --justify

Markdown 覆盖范围：
    标题（#/##/###…）、段落（软换行合并）、有序/无序/字母列表（保留原始编号与缩进层级）、
    管道表格（首行表头）、围栏代码块（等宽字体 + 浅底纹、保留空格）、引用块、分隔线（忽略）、
    行内 `代码` 与 **加粗**、行内与独立成行的图片。

图片支持：
    * 语法：![说明](路径) —— 兼容 ![说明](路径 "标题")、![说明](<带 空格.png>)、
      <img src="…" alt="…">；
    * 路径：相对路径按 Markdown 文件所在目录解析，也接受绝对路径、Windows 反斜杠、
      file:// 前缀与 URL 编码（%20）；
    * 排版：独立成行的图片居中显示，说明文字作居中图注（--no-caption 可关闭）；
      正文段落中的图片随文内联；表格单元格内的图片以单元格宽度为上限；
    * 缩放：只按比例缩小、不放大，上限为正文宽度（--image-width 可按百分比收紧）
      与正文高度的 62%，保证不越过页边距；
    * 容错：图片缺失或格式不受支持（如 EMF/SVG）时输出灰色占位文字并在终端告警，
      不中断转换。

目录页写入真正的 TOC 域；若本机装有 Word（pywin32 可用），默认调用 Word 更新目录页码，
否则退化为“打开文档时更新域”。

依赖：python-docx（生成文档）、pywin32（可选，仅用于刷新目录页码）；无需任何 Word 模板文件。
"""

from __future__ import annotations

import argparse
import datetime
import re
import sys
import unicodedata
from pathlib import Path
from urllib.parse import unquote

from docx import Document
from docx.image.image import Image as DocxImage
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

# ---------------------------------------------------------------- 版式常量
CN_FONT = "微软雅黑"
MONO_FONT = "Consolas"
BODY_PT = 10.5
CODE_PT = 8.0
TABLE_PT = 10.5
CODE_KEEP_LINES = 16
IMAGE_HEIGHT_RATIO = 0.62
IMAGE_PADDING_TWIPS = 240
CAPTION_PT = 9.0
IMAGE_MISSING_TEXT = "［图片缺失：%s］"

COLOR_TEXT = "2B3A4A"
COLOR_HEAD = "1F4E79"
COLOR_MUTED = "5B6B7B"
COLOR_CODE = "B03A2E"
BORDER_COLOR = "C9D4DE"
SHADE_HEAD = "EAF1F8"
SHADE_CODE = "F7F9FB"
SHADE_ZEBRA = "F7FAFD"

TOC_LEVELS = "2-3"

# 页面与样式（原本由模板文件携带，现由脚本现场生成）
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGIN_TOP_CM = 2.5
MARGIN_BOTTOM_CM = 2.0
MARGIN_LEFT_CM = 2.5
MARGIN_RIGHT_CM = 2.5
FOOTER_DISTANCE_CM = 1.5
BODY_LINE_SPACING = 1.21
BODY_SPACE_AFTER_PT = 10.0
HEADING_LINE_SPACING = 1.0625
COLOR_LEVEL4 = "4F81BD"
HEADING_SPECS = (
    # (样式名, 字号, 段前, 段后, 行距, 颜色, 对齐)；行距为 None 时跟随 --line-spacing
    ("Heading 1", 28.0, 9.0, 0.0, HEADING_LINE_SPACING, COLOR_HEAD, "center"),
    ("Heading 2", 15.5, 10.0, 0.0, HEADING_LINE_SPACING, COLOR_HEAD, "left"),
    ("Heading 3", 12.5, 9.0, 0.0, HEADING_LINE_SPACING, COLOR_HEAD, "left"),
    ("Heading 4", 10.5, 10.0, 0.0, None, COLOR_LEVEL4, "left"),
    ("Heading 5", 10.5, 10.0, 0.0, None, COLOR_LEVEL4, "left"),
    ("Heading 6", 10.5, 10.0, 0.0, None, COLOR_LEVEL4, "left"),
)
TOC_SPECS = (
    # (样式名, 左缩进 cm)；行距固定 13pt、段距为 0，否则目录会被 Normal 的段后距撑长
    ("TOC 1", 0.0),
    ("TOC 2", 0.74),
    ("TOC 3", 1.48),
)
TOC_LINE_PT = 13.0

# 可选纸张：键名小写，值为 (显示名, 宽 cm, 高 cm)
PAGE_SIZES = {
    "a4": ("A4", 21.0, 29.7),
    "a3": ("A3", 29.7, 42.0),
    "letter": ("Letter", 21.59, 27.94),
}
DEFAULT_MARGINS = (MARGIN_TOP_CM, MARGIN_BOTTOM_CM, MARGIN_LEFT_CM, MARGIN_RIGHT_CM)

# 封面主标题的排版宽度：默认版心（16 cm）下标定为 430 pt，换纸张/边距时按版心等比缩放
COVER_TITLE_WIDTH_PT = 430.0
COVER_REF_WIDTH_CM = 16.0


def parse_margin(text):
    """解析 --margin：'2.5' / '2.5,2' / '2.5,2,2.5,2.5' → (上, 下, 左, 右)，单位 cm。"""
    pieces = [piece.strip() for piece in str(text).replace("，", ",").split(",") if piece.strip()]
    try:
        values = [float(piece) for piece in pieces]
    except ValueError:
        raise argparse.ArgumentTypeError("边距只接受数字，形如 2.5,2,2.5,2.5（单位 cm）")
    if len(values) == 1:
        values = values * 4
    elif len(values) == 2:
        values = [values[0], values[0], values[1], values[1]]
    elif len(values) != 4:
        raise argparse.ArgumentTypeError("边距需给出 1、2 或 4 个数值，形如 2.5,2,2.5,2.5")
    if any(value < 0 for value in values):
        raise argparse.ArgumentTypeError("边距不能为负数")
    return tuple(values)


def parse_page_size(text):
    """校验 --page-size，返回 PAGE_SIZES 的键名。"""
    key = str(text).strip().lower().replace(" ", "")
    if key not in PAGE_SIZES:
        raise argparse.ArgumentTypeError("页面尺寸只支持 A4 / A3 / Letter")
    return key


class Layout:
    """页面与正文排版参数。

    缺省值与原先的硬编码常量完全相同，因此不带覆盖参数时输出与旧版一致；命令行
    覆盖后由 configure_page/configure_styles 写回分节与样式，text_width_emu、
    表格列宽、封面标题宽度与图片最大宽度都从分节重新推导，无需手工同步。
    """

    def __init__(self, page_width_cm=PAGE_WIDTH_CM, page_height_cm=PAGE_HEIGHT_CM,
                 margins=DEFAULT_MARGINS, body_pt=BODY_PT, line_spacing=BODY_LINE_SPACING,
                 justify=False, page_name="A4"):
        self.page_width_cm = page_width_cm
        self.page_height_cm = page_height_cm
        self.margins = tuple(margins)          # (上, 下, 左, 右) cm
        self.body_pt = body_pt
        self.line_spacing = line_spacing
        self.justify = justify
        self.page_name = page_name

    @property
    def text_width_cm(self):
        """版心宽（cm）。"""
        return self.page_width_cm - self.margins[2] - self.margins[3]

    @property
    def text_height_cm(self):
        """版心高（cm）。"""
        return self.page_height_cm - self.margins[0] - self.margins[1]

    @property
    def is_default(self):
        """是否完全沿用旧版默认版式。"""
        return (self.page_width_cm == PAGE_WIDTH_CM and self.page_height_cm == PAGE_HEIGHT_CM
                and self.margins == DEFAULT_MARGINS and self.body_pt == BODY_PT
                and self.line_spacing == BODY_LINE_SPACING and not self.justify)

    def describe(self):
        return ("%s %.2f×%.2f cm，边距 上%.2f 下%.2f 左%.2f 右%.2f，版心 %.2f×%.2f cm，"
                "正文 %.1fpt、行距 %.2f%s"
                % (self.page_name, self.page_width_cm, self.page_height_cm,
                   self.margins[0], self.margins[1], self.margins[2], self.margins[3],
                   self.text_width_cm, self.text_height_cm,
                   self.body_pt, self.line_spacing,
                   "、两端对齐" if self.justify else ""))

IMAGE_TAG_RE = re.compile(r"<img\b[^>]*>", re.IGNORECASE)
IMAGE_SRC_RE = re.compile(r"""src\s*=\s*["']([^"']+)["']""", re.IGNORECASE)
IMAGE_ALT_RE = re.compile(r"""alt\s*=\s*["']([^"']*)["']""", re.IGNORECASE)
IMAGE_TITLE_RE = re.compile(r"""\s+(?:"[^"]*"|'[^']*'|\([^)]*\))\s*$""")
HR_RE = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
TABLE_SEP_RE = re.compile(r"^:?-{2,}:?$")
LIST_RE = re.compile(r"^(\s*)([-*+]|\d+[.)]|[A-Za-z][.)])\s+(.*)$")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*$")

# OOXML 子元素顺序（Word 对 tcPr / pPr 的子元素顺序敏感）
TCPR_SEQ = (
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark",
)
PPR_SEQ = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
)
SECTPR_SEQ = (
    "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar", "w:paperSrc",
    "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt",
    "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi",
    "w:rtlGutter", "w:docGrid", "w:printerSettings",
)


def put_child(parent, seq, element):
    """按 OOXML 规定顺序插入子元素：插在本元素之后应出现的第一个兄弟节点之前。"""
    tag = element.tag
    names = [qn(name) for name in seq]
    index = names.index(tag)
    parent.insert_element_before(element, *seq[index + 1:])
    return element


# ---------------------------------------------------------------- 基础工具
def display_width(text: str) -> int:
    """按终端列宽估算字符串宽度：全角字符算 2 列。"""
    return sum(2 if unicodedata.east_asian_width(ch) in "WFA" else 1 for ch in text)


def set_run_font(run, size=None, bold=None, color=None, mono=False):
    """统一设置中西文字体；mono=True 时 ASCII 用等宽字体、中文仍用正文字体。"""
    name = MONO_FONT if mono else CN_FONT
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:cs"), CN_FONT)
    rfonts.set(qn("w:eastAsia"), CN_FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def scan_image_target(text, index):
    """解析 ![](…) 括号内内容：支持 <带空格路径>、"带空格路径" 与可选标题，返回 (src, end)。"""
    total = len(text)
    while index < total and text[index] in " \t":
        index += 1
    if index < total and text[index] in "<\"'":
        closer = ">" if text[index] == "<" else text[index]
        stop = text.find(closer, index + 1)
        if stop == -1:
            return None
        src = text[index + 1:stop].strip()
        index = stop + 1
    else:
        depth, stop = 0, index
        while stop < total:
            char = text[stop]
            if char == "(":
                depth += 1
            elif char == ")":
                if depth == 0:
                    break
                depth -= 1
            stop += 1
        if stop >= total:
            return None
        src = IMAGE_TITLE_RE.sub("", text[index:stop]).strip()
        index = stop
    while index < total and text[index] in " \t":
        index += 1
    if index < total and text[index] in "\"'":
        closer = text[index]
        stop = text.find(closer, index + 1)
        if stop == -1:
            return None
        index = stop + 1
        while index < total and text[index] in " \t":
            index += 1
    if index >= total or text[index] != ")" or not src:
        return None
    return src, index + 1


def scan_image(text, start=0):
    """从 start 起查找行内图片，返回 (begin, end, alt, src)；找不到返回 None。"""
    index = text.find("![", start)
    while index != -1:
        bracket = text.find("]", index + 2)
        if bracket == -1:
            return None
        if bracket + 1 < len(text) and text[bracket + 1] == "(":
            target = scan_image_target(text, bracket + 2)
            if target is not None:
                return index, target[1], text[index + 2:bracket], target[0]
        index = text.find("![", index + 2)
    return None


def standalone_image(text):
    """整行只有一张图片时返回 (alt, src)，否则 None。"""
    found = scan_image(text, 0)
    if found is None:
        return None
    begin, end, alt, src = found
    if text[:begin].strip() or text[end:].strip():
        return None
    return alt, src


def iter_inline(text):
    """切分行内元素，产出 (kind, value)：text / bold / code / image（image 的 value 为 (alt, src)）。"""
    index, buffer, total = 0, "", len(text)
    while index < total:
        if text.startswith("![", index):
            found = scan_image(text, index)
            if found is not None and found[0] == index:
                if buffer:
                    yield "text", buffer
                    buffer = ""
                yield "image", (found[2], found[3])
                index = found[1]
                continue
        if text.startswith("**", index):
            stop = text.find("**", index + 2)
            if stop > index + 2:
                if buffer:
                    yield "text", buffer
                    buffer = ""
                yield "bold", text[index + 2:stop]
                index = stop + 2
                continue
        if text[index] == "`":
            stop = text.find("`", index + 1)
            if stop > index + 1:
                if buffer:
                    yield "text", buffer
                    buffer = ""
                yield "code", text[index + 1:stop]
                index = stop + 1
                continue
        buffer += text[index]
        index += 1
    if buffer:
        yield "text", buffer


def resolve_image(src, base_dir):
    """把 Markdown 图片地址解析为本地路径：兼容绝对路径、相对路径、file://、反斜杠与 URL 编码。"""
    raw = (src or "").strip()
    if not raw:
        return None
    if raw.lower().startswith("file:"):
        raw = raw[5:]
    raw = unquote(raw).replace("\\", "/")
    if re.match(r"^/[A-Za-z]:", raw):  # file:///D:/… 这类多一个前导斜杠
        raw = raw[1:]
    candidate = Path(raw)
    if not candidate.is_absolute():
        roots = [Path(base_dir)] if base_dir else []
        roots.append(Path.cwd())
        for root in roots:
            probe = (root / raw).resolve()
            if probe.is_file():
                return probe
        return (roots[0] / raw).resolve() if roots else candidate
    return candidate


def insert_image(paragraph, src, ctx, max_width=None):
    """把一张图片插入段落：只按比例缩小，保证不超可用宽度与高度；失败仅告警。"""
    path = resolve_image(src, getattr(ctx, "base_dir", None))
    if path is None or not path.is_file():
        print("  [警告] 找不到图片：%s（以 %s 为基准目录）"
              % (src, getattr(ctx, "base_dir", Path.cwd())))
        if ctx is not None:
            ctx.images_missing = getattr(ctx, "images_missing", 0) + 1
        return False
    try:
        image = DocxImage.from_file(str(path))
    except Exception as exc:  # 含 EMF/WMF/SVG 等 python-docx 不支持的格式
        print("  [警告] 无法读取图片 %s：%s" % (src, exc))
        if ctx is not None:
            ctx.images_missing = getattr(ctx, "images_missing", 0) + 1
        return False
    width, height = int(image.width or 0), int(image.height or 0)
    if width <= 0 or height <= 0:  # 部分格式没有 DPI 信息，按 96dpi 估算
        width = int((image.px_width or 1) * 9525)
        height = int((image.px_height or 1) * 9525)
    limit_width = max_width or getattr(ctx, "max_image_width", width)
    limit_height = getattr(ctx, "max_image_height", height)
    scale = min(1.0, limit_width / width, limit_height / height)
    run = paragraph.add_run()
    run.add_picture(str(path), width=Emu(int(width * scale)), height=Emu(int(height * scale)))
    if ctx is not None:
        ctx.images_ok = getattr(ctx, "images_ok", 0) + 1
    return True


def add_inline(paragraph, text, size=BODY_PT, bold_all=False, color=COLOR_TEXT,
               ctx=None, max_width=None):
    """把一段 Markdown 文本写入段落，处理行内 `代码`、**加粗** 与 ![图片](路径)。"""
    for kind, value in iter_inline(text):
        if kind == "image":
            alt, src = value
            if not insert_image(paragraph, src, ctx, max_width):
                run = paragraph.add_run(IMAGE_MISSING_TEXT % (alt or src))
                set_run_font(run, size - 0.5, color=COLOR_MUTED)
        elif kind == "bold":
            run = paragraph.add_run(value)
            set_run_font(run, size, bold=True, color=color)
        elif kind == "code":
            run = paragraph.add_run(value)
            set_run_font(run, size - 0.5, color=COLOR_CODE, mono=True)
        else:
            run = paragraph.add_run(value)
            set_run_font(run, size, bold=bold_all, color=color)


def body_size(ctx):
    """取本次转换的正文磅值（--body-pt）；未传入 ctx 时回落到默认值。"""
    return getattr(ctx, "body_pt", BODY_PT)


def add_image_block(doc, alt, src, cfg):
    """独立成行的图片：居中显示，alt 文本作为居中图注（--no-caption 可关闭）。"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fmt = para.paragraph_format
    fmt.space_before = Pt(6)
    with_caption = bool(alt) and cfg.caption
    fmt.space_after = Pt(3) if with_caption else Pt(9)
    fmt.keep_with_next = with_caption
    if not insert_image(para, src, cfg):
        run = para.add_run(IMAGE_MISSING_TEXT % src)
        set_run_font(run, CAPTION_PT, color=COLOR_MUTED)
        return para
    if not with_caption:
        return para
    caption = doc.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_fmt = caption.paragraph_format
    c_fmt.space_before = Pt(0)
    c_fmt.space_after = Pt(9)
    add_inline(caption, alt, size=CAPTION_PT, color=COLOR_MUTED, ctx=cfg)
    return caption


def add_spacer(doc, pt=6):
    """插入一个高度固定的空段落，用作表格/代码块后的间距。"""
    para = doc.add_paragraph()
    fmt = para.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(pt)
    return para


def shade_cell(cell, hex_color):
    """给表格单元格加底纹。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    put_child(tc_pr, TCPR_SEQ, shd)


def border_cell(cell, color=BORDER_COLOR, size=6):
    """给单元格四边加细线边框（逐单元格设置）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement("w:" + edge)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)
    put_child(tc_pr, TCPR_SEQ, borders)


def valign_cell(cell, value="center"):
    """单元格垂直居中。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), value)
    put_child(tc_pr, TCPR_SEQ, v_align)


def cell_padding(cell, top=100, left=120, bottom=100, right=120):
    """单元格内边距（twips）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        element = OxmlElement("w:" + edge)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")
        margins.append(element)
    put_child(tc_pr, TCPR_SEQ, margins)


def mark_header_row(row):
    """标记为标题行：跨页时自动重复，禁止行内断页，并与后续数据行同页。"""
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)
    keep_row_together(row)
    for cell in row.cells:
        for para in cell.paragraphs:
            para.paragraph_format.keep_with_next = True


def keep_row_together(row):
    """禁止该行内部跨页断开（行高超过一页时 Word 仍会断开）。"""
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.insert(0, cant_split)


def add_field(paragraph, instr, placeholder="", size=None, color=None):
    """在段落中插入一个域（如 PAGE / TOC），placeholder 为域结果占位文本。"""
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run = paragraph.add_run()
    run._r.append(begin)
    run2 = paragraph.add_run()
    run2._r.append(instr_el)
    run3 = paragraph.add_run()
    run3._r.append(separate)
    if placeholder:
        run4 = paragraph.add_run(placeholder)
        set_run_font(run4, size, color=color)
    run5 = paragraph.add_run()
    run5._r.append(end)


def fit_size(text, max_pt=28.0, min_pt=16.0, width_pt=430.0):
    """按可用宽度自动收缩封面主标题字号，避免折行。"""
    units = max(display_width(text), 1)
    size = width_pt * 2.0 / units
    return max(min_pt, min(max_pt, round(size * 2) / 2.0))


# ---------------------------------------------------------------- 文档骨架
def clear_theme_fonts(rfonts):
    """删除 rFonts 上的主题字体属性，否则 Word 会忽略显式指定的字体名。"""
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        key = qn("w:" + attr)
        if key in rfonts.attrib:
            del rfonts.attrib[key]


def set_style_font(style, size=None, bold=None, color=None, italic=False):
    """设置样式的中西文字体、字号、粗细与颜色。

    空白文档自带的 Heading 样式用的是主题字体（majorHAnsi）与主题色（accent1），
    主题属性优先级高于显式值，因此先把它们清掉再写固定值。
    """
    style.font.name = CN_FONT
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    clear_theme_fonts(rfonts)
    rfonts.set(qn("w:ascii"), CN_FONT)
    rfonts.set(qn("w:hAnsi"), CN_FONT)
    rfonts.set(qn("w:cs"), CN_FONT)
    rfonts.set(qn("w:eastAsia"), CN_FONT)
    for element in rpr.findall(qn("w:color")):  # 主题色同样会盖过固定色值
        rpr.remove(element)
    style.font.italic = italic
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = RGBColor.from_string(color)


def configure_page(doc, layout):
    """按 layout 设置纸张与页边距（缺省 A4、上下 2.5/2.0 cm、左右 2.5 cm）。"""
    section = doc.sections[0]
    section.page_width = Cm(layout.page_width_cm)
    section.page_height = Cm(layout.page_height_cm)
    section.top_margin = Cm(layout.margins[0])
    section.bottom_margin = Cm(layout.margins[1])
    section.left_margin = Cm(layout.margins[2])
    section.right_margin = Cm(layout.margins[3])
    section.footer_distance = Cm(FOOTER_DISTANCE_CM)


def ensure_toc_styles(doc, layout):
    """预建 TOC 1/2/3 样式，供 Word 生成目录时直接复用。

    Word 默认按 Normal 派生目录样式，而 Normal 带有段后距，条目一多就会把目录
    撑到第二页；这里改为固定行距（默认 13pt）、零段距，并按层级给出左缩进。
    """
    normal = doc.styles["Normal"]
    existing = [style.name for style in doc.styles]
    line_pt = max(TOC_LINE_PT, layout.body_pt + 2.5)  # 字号调大时同步放宽行距
    for name, indent in TOC_SPECS:
        if name in existing:
            style = doc.styles[name]
        else:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
        style.base_style = normal
        set_style_font(style, layout.body_pt, color=COLOR_TEXT)
        fmt = style.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = Pt(line_pt)
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if indent:
            fmt.left_indent = Cm(indent)


def configure_styles(doc, layout):
    """在空白文档上重建正文与标题样式，替代模板自带的 styles.xml。"""
    normal = doc.styles["Normal"]
    set_style_font(normal, layout.body_pt, color=COLOR_TEXT)
    fmt = normal.paragraph_format
    fmt.line_spacing = layout.line_spacing
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(BODY_SPACE_AFTER_PT)
    fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if layout.justify else WD_ALIGN_PARAGRAPH.LEFT

    for name, size, before, after, line, color, align in HEADING_SPECS:
        style = doc.styles[name]
        set_style_font(style, size, bold=True, color=color)
        h_fmt = style.paragraph_format
        h_fmt.line_spacing = layout.line_spacing if line is None else line
        h_fmt.space_before = Pt(before)
        h_fmt.space_after = Pt(after)
        h_fmt.keep_with_next = True
        h_fmt.alignment = (WD_ALIGN_PARAGRAPH.CENTER if align == "center"
                           else WD_ALIGN_PARAGRAPH.LEFT)

    if "List Bullet" in [style.name for style in doc.styles]:
        bullet = doc.styles["List Bullet"]
        bullet.base_style = normal
        set_style_font(bullet, layout.body_pt, color=COLOR_TEXT)

    ensure_toc_styles(doc, layout)


def new_document(layout=None):
    """按布局参数建立空白文档，并把页面与样式配好。"""
    layout = layout if layout is not None else Layout()
    doc = Document()
    configure_page(doc, layout)
    configure_styles(doc, layout)
    return doc


def text_width_twips(doc):
    """正文可用宽度（twips）。"""
    section = doc.sections[0]
    usable = section.page_width - section.left_margin - section.right_margin
    return int(usable / 635)  # 1 twip = 635 EMU


def text_width_emu(doc):
    """正文可用宽度（EMU）。"""
    return text_width_twips(doc) * 635


def text_width_pt(doc):
    """正文可用宽度（磅）。"""
    return text_width_twips(doc) / 20.0


def text_height_emu(doc):
    """正文可用高度（EMU），用于限制图片高度不越过上下边距。"""
    section = doc.sections[0]
    return int(section.page_height - section.top_margin - section.bottom_margin)


def style_heading(doc, text, level):
    """按层级输出标题：章=Heading 2，节=Heading 3。"""
    style_name = "Heading %d" % min(max(level, 2), 4)
    if style_name not in [s.name for s in doc.styles]:
        style_name = "Heading 3"
    para = doc.add_paragraph(style=style_name)
    for kind, value in iter_inline(text):
        if kind == "code":
            run = para.add_run(value)
            set_run_font(run, color=COLOR_HEAD, mono=True)
        elif kind == "bold":
            run = para.add_run(value)
            set_run_font(run, color=COLOR_HEAD, bold=True)
        elif kind == "image":  # 标题行不放图，退化为图片说明文字
            run = para.add_run(value[0] or value[1])
            set_run_font(run, color=COLOR_HEAD)
        else:
            run = para.add_run(value)
            set_run_font(run, color=COLOR_HEAD)
    return para


def style_lead(doc, text, size, bold=True, color=COLOR_HEAD, border=False, align=None, ctx=None):
    """「版本历史」「目　　录」这类小节标题的样式。"""
    para = doc.add_paragraph()
    if align is not None:
        para.alignment = align
    fmt = para.paragraph_format
    fmt.space_after = Pt(9)
    if border:
        p_bdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "9")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), SHADE_HEAD)
        p_bdr.append(bottom)
        put_child(para._p.get_or_add_pPr(), PPR_SEQ, p_bdr)
    add_inline(para, text, size=size, bold_all=bold, color=color, ctx=ctx)
    return para


# ---------------------------------------------------------------- 各类块渲染
def add_paragraph_block(doc, text, ctx=None, indent=0):
    para = doc.add_paragraph()
    if indent:
        para.paragraph_format.left_indent = Cm(indent)
    add_inline(para, text, size=body_size(ctx), ctx=ctx)
    return para


def add_quote_block(doc, text, ctx=None):
    para = doc.add_paragraph()
    fmt = para.paragraph_format
    fmt.left_indent = Cm(0.6)
    add_inline(para, text, size=body_size(ctx), ctx=ctx)
    for run in para.runs:
        run.italic = True
    return para


def marker_indent(marker, level=0, size=BODY_PT):
    """按编号实测宽度算悬挂缩进（厘米）。

    编号后用不换行空格分隔而不靠制表符：Word 的两端对齐在个别长行会把制表符压成零宽，
    导致“8.有”这类编号与正文粘连；不换行空格不参与两端对齐的字间距分配，宽度恒定。
    字形宽度取自微软雅黑 10.5pt 的实测值（数字 0.589em、字母 0.62em、句点 0.243em、空格 0.298em）。
    """
    total = 0.0
    for char in marker:
        if char.isdigit():
            total += 0.589
        elif char.isalpha():
            total += 0.62
        elif char in ".)：":
            total += 0.243
        else:
            total += 0.55
    total += 0.298  # 编号与正文之间的空格
    return (0.635 + 0.6 * level) + total * size / 28.35


def add_list_block(doc, marker, ordered, level, text, ctx=None):
    """有序/字母列表使用原始编号（保留 Markdown 中的序号），无序列表使用内置 List Bullet 样式。"""
    para = doc.add_paragraph()
    fmt = para.paragraph_format
    base = 0.635 + 0.6 * level
    size = body_size(ctx)
    if ordered:
        indent = marker_indent(marker, level, size)
        fmt.left_indent = Cm(indent)
        fmt.first_line_indent = Cm(base - indent)
        add_inline(para, marker + "\u00A0" + text, size=size, ctx=ctx)
    else:
        para.style = doc.styles["List Bullet"]
        if level:
            fmt.left_indent = Cm(base + 0.55)
            fmt.first_line_indent = Cm(-0.55)
        add_inline(para, text, size=size, ctx=ctx)
    return para


def plan_widths(rows, total_twips, size=TABLE_PT):
    """规划表格列宽（twips）。

    want：每列最长内容的单行宽度，用 want**0.75 作为权重，兼顾内容宽度与列间均衡；
    hard：每列的最小行宽度 —— 取该列最长不可断开 ASCII 串的宽度（不超过表宽 1/3），
          保证路径、标识符不会被拦腰截断；同时不低于 42pt。
    先按权重分配，把不够 hard 的列顶到 hard，剩余宽度在其余列中再分配。
    """
    cols = len(rows[0])
    ascii_ratio = 0.58

    def text_pt(text):
        return sum(size * (1.0 if unicodedata.east_asian_width(ch) in "WFA" else ascii_ratio)
                   for ch in text)

    total_pt = total_twips / 20.0
    want, hard = [], []
    for index in range(cols):
        want.append(max(24.0, max(text_pt(row[index]) for row in rows)))
        longest_token = 0.0
        for row in rows:
            for token in re.findall(r"[^\s\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]+", row[index]):
                longest_token = max(longest_token, text_pt(token))
        hard.append(min(max(longest_token + 8.0, 44.0), total_pt * 0.28))

    weights = [value ** 0.75 for value in want]
    widths = [total_pt * weight / sum(weights) for weight in weights]
    # 先撑到每列的下限，再在“可压缩空间”内收放，保证不破下限
    widths = [max(width, floor) for width, floor in zip(widths, hard)]
    for _ in range(6):
        diff = total_pt - sum(widths)
        if abs(diff) < 0.5:
            break
        if diff < 0:
            slack = [max(0.0, width - floor) for width, floor in zip(widths, hard)]
            if sum(slack) <= 0:
                break
            widths = [width - (-diff) * space / sum(slack)
                      for width, space in zip(widths, slack)]
        else:
            widths = [width + diff * weight / sum(weights)
                      for width, weight in zip(widths, weights)]

    # 把残差压到最宽的列，保证表格总宽等于正文宽度
    widest = max(range(cols), key=lambda i: widths[i])
    widths[widest] = max(1.0, widths[widest] + (total_pt - sum(widths)))
    return [int(round(width * 20)) for width in widths]


def add_table_block(doc, rows, ctx=None):
    """渲染管道表格：统一配色边框、表头底纹加粗、跨页重复表头。"""
    cols = len(rows[0])
    widths = plan_widths(rows, text_width_twips(doc))

    table = doc.add_table(rows=len(rows), cols=cols)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, width in enumerate(widths):
        table.columns[index].width = Pt(width / 20.0)

    for r_index, cells in enumerate(rows):
        for c_index, value in enumerate(cells):
            cell = table.cell(r_index, c_index)
            cell.width = Pt(widths[c_index] / 20.0)
            border_cell(cell)
            valign_cell(cell)
            para = cell.paragraphs[0]
            fmt = para.paragraph_format
            fmt.space_before = Pt(3)
            fmt.space_after = Pt(3)
            fmt.line_spacing = 1.0
            if getattr(ctx, "justify", False):  # 两端对齐会把单元格里的短文本拉散，表格恒左对齐
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # 单元格内的图片以扣除内边距后的列宽为上限
            cell_width = max(72 * 635, (widths[c_index] - IMAGE_PADDING_TWIPS) * 635)
            add_inline(para, value, size=TABLE_PT,
                       bold_all=(r_index == 0),
                       color=COLOR_HEAD if r_index == 0 else COLOR_TEXT,
                       ctx=ctx, max_width=cell_width)
            if r_index == 0:
                shade_cell(cell, SHADE_HEAD)
            elif r_index % 2 == 0:
                shade_cell(cell, SHADE_ZEBRA)
    mark_header_row(table.rows[0])
    # 数据行也不跨页断开：宁可整行移到下一页，避免单元格被“腰斩”。
    for row in table.rows[1:]:
        keep_row_together(row)
    add_spacer(doc, 6)


def add_code_block(doc, lines, ctx=None):
    """代码块放进单格表格：等宽字体 + 浅底纹，逐行渲染以保留原始空格。"""
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Pt(text_width_twips(doc) / 20.0)
    cell = table.cell(0, 0)
    cell.width = Pt(text_width_twips(doc) / 20.0)
    border_cell(cell)
    shade_cell(cell, SHADE_CODE)
    cell_padding(cell)

    for index, line in enumerate(lines):
        para = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        fmt = para.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        if getattr(ctx, "justify", False):  # 代码行不能被两端对齐拉开空格
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = para.add_run(line if line else " ")
        set_run_font(run, CODE_PT, color="1A1A1A", mono=True)
    # 短代码块整块不拆；长清单允许跨页续排，避免整页留白。
    if len(lines) <= CODE_KEEP_LINES:
        keep_row_together(table.rows[0])
    add_spacer(doc, 6)


# ---------------------------------------------------------------- 前置页面
def add_cover(doc, title, meta_lines, cfg):
    """封面页：主标题 + Markdown 首部信息行 + 版本/日期/密级。"""
    if cfg.kicker:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(112.5)
        run = para.add_run(cfg.kicker)
        set_run_font(run, 11, color=COLOR_MUTED)

    title_para = doc.add_paragraph(style="Heading 1")
    title_para.paragraph_format.space_before = Pt(0 if cfg.kicker else 150)
    title_para.paragraph_format.space_after = Pt(18)
    size = fit_size(title, width_pt=COVER_TITLE_WIDTH_PT *
                     (cfg.layout.text_width_cm / COVER_REF_WIDTH_CM))
    run = title_para.add_run(title)
    set_run_font(run, size, bold=True, color=COLOR_HEAD)

    for line in meta_lines:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(4)
        add_inline(para, line, size=11, color=COLOR_MUTED, ctx=cfg)

    for index, line in enumerate(("版本 %s" % cfg.version, "日期 %s" % cfg.date, "密级：%s" % cfg.secret)):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(2)
        if index == 0:
            para.paragraph_format.space_before = Pt(84)
        run = para.add_run(line)
        set_run_font(run, 10, color=COLOR_MUTED)


def add_version_history(doc, cfg):
    """版本历史页：与前文同构的表格 + 修订规则说明。"""
    style_lead(doc, "版本历史", 15.5, border=True, ctx=cfg)
    rows = [
        ["版本号", "日期", "著作者", "修改内容"],
        [cfg.version, cfg.date, cfg.author, "初版：由 %s 转换生成" % cfg.source_name],
    ]
    add_table_block(doc, rows, ctx=cfg)
    note = ("后续修订规则：源 Markdown 内容更新并经确认后，本文档递增版本号并在上表登记变更行；"
            "未发生对应源内容变更的编辑性修改不单独出版本。")
    add_paragraph_block(doc, note, ctx=cfg)


def add_toc_page(doc, cfg=None):
    """目录页：写入 TOC 域，由 Word 填充条目与页码。"""
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading.paragraph_format.space_after = Pt(18)
    run = heading.add_run("目　　录")
    set_run_font(run, 15.5, bold=True, color=COLOR_HEAD)

    para = doc.add_paragraph()
    add_field(para, ' TOC \\o "%s" \\h \\z \\u ' % TOC_LEVELS,
              placeholder="（打开文档后右键“更新域”生成目录条目与页码）",
              size=body_size(cfg), color=COLOR_MUTED)


# ---------------------------------------------------------------- 分节收尾
def strip_header_footer(section):
    """去掉分节的页眉页脚引用：封面/版本历史/目录三页不显示页码。"""
    sect_pr = section._sectPr
    for tag in ("w:headerReference", "w:footerReference"):
        for element in sect_pr.findall(qn(tag)):
            sect_pr.remove(element)


def ensure_page_number_footer(section, doc):
    """正文分节页脚：居中 PAGE 域；并确保页码从 1 重新编号。"""
    footer = section.footer
    footer.is_linked_to_previous = False
    para = footer.paragraphs[0]
    for run in list(para.runs):
        run._r.getparent().remove(run._r)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(para, " PAGE \\* MERGEFORMAT ", "1", size=9, color=COLOR_MUTED)
    for run in para.runs:
        set_run_font(run, 9, color=COLOR_MUTED)

    sect_pr = section._sectPr
    for element in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(element)
    pg_num = OxmlElement("w:pgNumType")
    pg_num.set(qn("w:start"), "1")
    put_child(sect_pr, SECTPR_SEQ, pg_num)


def finalize_sections(doc):
    """分节收尾：前置页去页码，最后一节带页码。"""
    sections = list(doc.sections)
    for section in sections[:-1]:
        strip_header_footer(section)
    ensure_page_number_footer(sections[-1], doc)


def set_update_fields(doc, enabled=True):
    """控制“打开文档时更新域”，保证没有 Word 时目录也能刷新。"""
    settings = doc.settings.element
    existing = settings.find(qn("w:updateFields"))
    if not enabled:
        if existing is not None:
            settings.remove(existing)
        return
    if existing is None:
        element = OxmlElement("w:updateFields")
        element.set(qn("w:val"), "true")
        settings.append(element)


# ---------------------------------------------------------------- Markdown 解析
def split_row(line):
    """拆分 | a | b | 形式的表格行。"""
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split("|")]


def is_separator_row(cells):
    return all(TABLE_SEP_RE.match(cell) for cell in cells if cell)


def is_block_start(line):
    """判断当前行是否开启一个新块（用于段落收集时断句）。"""
    stripped = line.strip()
    if not stripped:
        return True
    if stripped.startswith("```") or stripped.startswith("~~~"):
        return True
    if stripped.startswith("|") and stripped.count("|") >= 2:
        return True
    if stripped.startswith("#") and HEADING_RE.match(stripped):
        return True
    if HR_RE.match(stripped) or stripped.startswith(">"):
        return True
    if standalone_image(stripped) is not None:
        return True
    return bool(LIST_RE.match(line))


def join_soft_lines(buffer):
    """合并段落内的软换行：中文之间不加空格，英文单词之间补一个空格。"""
    out = ""
    for line in buffer:
        line = line.strip()
        if not out:
            out = line
        elif out[-1].isascii() and out[-1].isalnum() and line[:1].isascii() and line[:1].isalnum():
            out += " " + line
        else:
            out += line
    return out


def parse_blocks(lines):
    """把 Markdown 解析为块列表。"""
    blocks = []
    index, total = 0, len(lines)
    while index < total:
        line = lines[index].rstrip()
        stripped = line.strip()

        if stripped.startswith("```") or stripped.startswith("~~~"):
            fence = stripped[:3]
            index += 1
            body = []
            while index < total and not lines[index].strip().startswith(fence):
                body.append(lines[index].rstrip())
                index += 1
            index += 1
            blocks.append({"type": "code", "lines": body})
            continue

        if stripped.startswith("|") and stripped.count("|") >= 2:
            rows = []
            while index < total and lines[index].strip().startswith("|"):
                cells = split_row(lines[index])
                if not is_separator_row(cells):
                    rows.append(cells)
                index += 1
            if rows:
                width = max(len(row) for row in rows)
                rows = [row + [""] * (width - len(row)) for row in rows]
                blocks.append({"type": "table", "rows": rows})
            continue

        heading = HEADING_RE.match(stripped)
        if heading:
            blocks.append({"type": "heading",
                           "level": len(heading.group(1)),
                           "text": heading.group(2).strip()})
            index += 1
            continue

        if HR_RE.match(stripped):
            blocks.append({"type": "hr"})
            index += 1
            continue

        if stripped.startswith(">"):
            blocks.append({"type": "quote", "text": stripped.lstrip(">").strip()})
            index += 1
            continue

        item = LIST_RE.match(line)
        if item:
            indent = len(item.group(1).replace("\t", "    "))
            marker = item.group(2)
            blocks.append({"type": "list",
                           "marker": marker,
                           "ordered": not marker[0] in "-*+",
                           "level": min(indent // 3, 3),
                           "text": item.group(3).strip()})
            index += 1
            continue

        image = standalone_image(stripped)
        if image is not None:
            blocks.append({"type": "image", "alt": image[0], "src": image[1]})
            index += 1
            continue

        buffer = []
        while index < total and not is_block_start(lines[index]):
            buffer.append(lines[index].rstrip())
            index += 1
        if buffer:
            blocks.append({"type": "para", "lines": buffer})
        else:  # 兜底，避免死循环
            index += 1
    return blocks


def replace_html_image(match):
    """把单个 <img> 标签改写为 Markdown 图片语法；无 src 时原样返回。"""
    tag = match.group(0)
    src = IMAGE_SRC_RE.search(tag)
    if not src:
        return tag
    alt = IMAGE_ALT_RE.search(tag)
    return "![%s](%s)" % (alt.group(1) if alt else "", src.group(1).strip())


def normalize_html_images(lines):
    """预先把 <img src="…"> 统一成 ![](…)，围栏代码块内的内容保持原样。"""
    out, fence = [], None
    for line in lines:
        stripped = line.strip()
        if fence:
            if stripped.startswith(fence):
                fence = None
            out.append(line)
            continue
        if stripped.startswith("```") or stripped.startswith("~~~"):
            fence = stripped[:3]
            out.append(line)
            continue
        out.append(IMAGE_TAG_RE.sub(replace_html_image, line))
    return out


# ---------------------------------------------------------------- 转换主流程
def split_front(blocks, fallback_title):
    """拆出封面用的标题与信息行，返回 (title, meta_lines, 正文起始下标)。"""
    title = None
    meta_lines = []
    index = 0
    if blocks and blocks[0]["type"] == "heading" and blocks[0]["level"] == 1:
        title = blocks[0]["text"]
        index = 1
    if index < len(blocks) and blocks[index]["type"] == "para":
        meta_lines = [line.strip() for line in blocks[index]["lines"] if line.strip()]
        index += 1
    while index < len(blocks) and blocks[index]["type"] == "hr":
        index += 1
    return title or fallback_title, meta_lines, index


def render_body(doc, blocks, cfg):
    """按块类型渲染正文；表格/代码/图片块前一句设为“与下段同页”，避免标题行落单。"""
    last_para = None
    for block in blocks:
        kind = block["type"]
        if kind == "heading":
            last_para = style_heading(doc, block["text"], block["level"])
        elif kind == "para":
            last_para = add_paragraph_block(doc, join_soft_lines(block["lines"]), ctx=cfg)
        elif kind == "list":
            last_para = add_list_block(doc, block["marker"], block["ordered"],
                                       block["level"], block["text"], ctx=cfg)
        elif kind in ("table", "code", "image"):
            if last_para is not None:
                last_para.paragraph_format.keep_with_next = True
            if kind == "table":
                add_table_block(doc, block["rows"], ctx=cfg)
            elif kind == "code":
                add_code_block(doc, block["lines"], ctx=cfg)
            else:
                add_image_block(doc, block["alt"], block["src"], cfg)
            last_para = None
        elif kind == "quote":
            last_para = add_quote_block(doc, block["text"], ctx=cfg)
        # hr：正文不使用分隔线，直接忽略


def refresh_toc_via_word(docx_path: Path) -> bool:
    """调用本机 Word 更新目录域，使页码真实落地；不可用时返回 False。"""
    try:
        import win32com.client as win32
    except ImportError:
        return False
    app = None
    try:
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        document = app.Documents.Open(str(docx_path), ReadOnly=False, AddToRecentFiles=False)
        document.Fields.Update()
        for index in range(1, document.TablesOfContents.Count + 1):
            document.TablesOfContents(index).Update()
        document.Fields.Update()
        document.Save()
        document.Close(SaveChanges=0)
        return True
    except Exception as exc:  # 任何 COM 异常都不影响文档本身
        print("  [提示] Word 更新目录失败：%s" % exc)
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def build(md_path: Path, docx_path: Path, cfg) -> Path:
    lines = normalize_html_images(md_path.read_text(encoding="utf-8").splitlines())
    blocks = parse_blocks(lines)
    title, meta_lines, body_start = split_front(blocks, md_path.stem)

    doc = new_document(cfg.layout)
    # 图片尺寸上限按本文档的版心计算，全文共用；换纸张/边距后此处自动重算
    cfg.max_image_width = int(text_width_emu(doc) * cfg.image_width_pct)
    cfg.max_image_height = int(text_height_emu(doc) * IMAGE_HEIGHT_RATIO)

    if cfg.cover:
        add_cover(doc, title, meta_lines, cfg)
        doc.add_section(WD_SECTION.NEW_PAGE)
    if cfg.history:
        add_version_history(doc, cfg)
        doc.add_section(WD_SECTION.NEW_PAGE)
    if cfg.toc:
        add_toc_page(doc, cfg)
        doc.add_section(WD_SECTION.NEW_PAGE)

    render_body(doc, blocks[body_start:], cfg)
    finalize_sections(doc)

    props = doc.core_properties
    props.title = title
    props.subject = md_path.name
    props.comments = "由 local3.py 转换生成（版式由脚本现场生成，不依赖外部模板）"
    if cfg.author:
        props.author = cfg.author

    set_update_fields(doc, enabled=True)
    docx_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(docx_path))

    if cfg.toc and cfg.update_toc and refresh_toc_via_word(docx_path):
        doc = Document(str(docx_path))
        set_update_fields(doc, enabled=False)
        doc.save(str(docx_path))
    return docx_path


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="把 Markdown 转换为带封面/版本历史/目录的 Word 文档（版式由脚本生成，无需模板文件）")
    parser.add_argument("input", help="输入的 Markdown 文件")
    parser.add_argument("output", nargs="?", help="输出的 docx（缺省与输入同名）")
    parser.add_argument("-t", "--template", default=None,
                        help="已废弃：版式改由脚本生成，此参数被忽略（仅为兼容旧命令行保留）")
    parser.add_argument("--version", default="V1.0", help="封面版本号，缺省 V1.0")
    parser.add_argument("--date", default=None, help="封面日期，缺省为当天（YYYY-MM-DD）")
    parser.add_argument("--author", default="（待填）", help="版本历史中的著作者")
    parser.add_argument("--secret", default="内部", help="封面密级，缺省 内部")
    parser.add_argument("--kicker", default="", help="封面主标题上方的灰色小字（可略）")
    parser.add_argument("--no-cover", action="store_true", help="不生成封面页")
    parser.add_argument("--no-history", action="store_true", help="不生成版本历史页")
    parser.add_argument("--no-toc", action="store_true", help="不生成目录页")
    parser.add_argument("--no-word", action="store_true", help="不调用本机 Word 更新目录页码")
    parser.add_argument("--image-width", type=float, default=100.0, metavar="PCT",
                        help="图片最大宽度占版心的百分比，缺省 100（只缩小不放大）")
    parser.add_argument("--image-dir", default=None,
                        help="图片相对路径的基准目录，缺省为 Markdown 文件所在目录")
    parser.add_argument("--no-caption", action="store_true",
                        help="不把图片的 alt 文字渲染为居中图注")
    parser.add_argument("--page-size", type=parse_page_size, default="a4", metavar="A4|A3|Letter",
                        help="纸张尺寸，缺省 A4；也可写 a4/a3/letter")
    parser.add_argument("--margin", type=parse_margin, default=None, metavar="上,下,左,右",
                        help="页边距（cm），缺省 2.5,2.0,2.5,2.5；可只给 1 个或 2 个值")
    parser.add_argument("--body-pt", type=float, default=BODY_PT, metavar="PT",
                        help="正文字号（磅），缺省 10.5")
    parser.add_argument("--line-spacing", type=float, default=BODY_LINE_SPACING, metavar="X",
                        help="正文行距倍数，缺省 1.21；Heading 4 的行距跟随此值")
    parser.add_argument("--justify", action="store_true",
                        help="正文两端对齐（缺省左对齐）")
    return parser.parse_args(argv)


def make_layout(args):
    """由命令行参数构造 Layout；校验失败返回 (None, 错误信息)。"""
    if not 6.0 <= args.body_pt <= 36.0:
        return None, "正文字号需在 6–36 pt 之间：%s" % args.body_pt
    if not 0.8 <= args.line_spacing <= 4.0:
        return None, "行距倍数需在 0.8–4.0 之间：%s" % args.line_spacing
    name, width, height = PAGE_SIZES[args.page_size]
    margins = args.margin if args.margin else DEFAULT_MARGINS
    layout = Layout(width, height, margins, args.body_pt, args.line_spacing,
                    args.justify, name)
    if layout.text_width_cm < 5.0 or layout.text_height_cm < 5.0:
        return None, ("页面与边距组合不可用：版心仅 %.2f×%.2f cm"
                      % (layout.text_width_cm, layout.text_height_cm))
    return layout, None


def main(argv=None):
    args = parse_args(argv)
    md_path = Path(args.input).expanduser().resolve()
    if not md_path.is_file():
        print("找不到输入文件：%s" % md_path)
        return 2
    docx_path = Path(args.output).expanduser().resolve() if args.output else md_path.with_suffix(".docx")

    if args.template:
        print("  [提示] 已忽略 -t/--template：local3.py 不再读取外部模板，版式由脚本现场生成")

    layout, error = make_layout(args)
    if layout is None:
        print("  [错误] %s" % error)
        return 2

    args.source_name = md_path.name
    args.date = args.date or datetime.date.today().isoformat()
    args.cover = not args.no_cover
    args.history = not args.no_history
    args.toc = not args.no_toc
    args.update_toc = not args.no_word
    args.caption = not args.no_caption
    args.image_width_pct = min(1.0, max(0.1, args.image_width / 100.0))
    args.base_dir = Path(args.image_dir).expanduser().resolve() if args.image_dir else md_path.parent
    args.layout = layout
    args.body_pt = layout.body_pt          # 正文磅值随 Layout 下发到各渲染函数
    if not layout.is_default:
        print("  [版式] %s" % layout.describe())

    out = build(md_path, docx_path, args)
    ok = getattr(args, "images_ok", 0)
    missing = getattr(args, "images_missing", 0)
    if ok or missing:
        print("图片：已插入 %d 张，缺失或不支持 %d 张" % (ok, missing))
    print("已生成：%s（%.1f KB）" % (out, out.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
