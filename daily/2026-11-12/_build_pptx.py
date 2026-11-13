# -*- coding: utf-8 -*-
"""Newspaper-style 16:9 briefing. Visual system shared with briefing.html."""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _stories import DATE, DATE_CN, FOLIO, ISSUE, POLITICS, TECH, WEEKDAY  # noqa: E402

INK = RGBColor(0x1A, 0x1A, 0x1A)
PAPER = RGBColor(0xF7, 0xF4, 0xEF)
CRIMSON = RGBColor(0xC4, 0x1E, 0x3A)
SLATE = RGBColor(0x4A, 0x55, 0x68)
CARD = RGBColor(0xEF, 0xEB, 0xE3)
RULE = RGBColor(0xC9, 0xC2, 0xB6)
CREAM = RGBColor(0xFB, 0xF8, 0xF3)
MUTED = RGBColor(0x6B, 0x63, 0x58)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SERIF_L, SERIF_E = "Cambria", "Noto Serif CJK SC"
SANS_L, SANS_E = "Calibri", "Noto Sans CJK SC"

W, H = Inches(13.333333), Inches(7.5)
M = 0.55


def rgb_hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def fill_solid(shape, color: RGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def fill_line(shape, fill: RGBColor, line: RGBColor, pt: float = 0.75):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(pt)


def rect(slide, l, t, w, h, color, line=None, line_pt=0.75):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    if line is None:
        fill_solid(sh, color)
    else:
        fill_line(sh, color, line, line_pt)
    sh.shadow.inherit = False
    return sh


def rule(slide, l, t, w, color=INK, thick=0.012):
    return rect(slide, l, t, w, thick, color)


def apply_font(run, latin, ea, size, bold, color, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = latin
    rPr = run._r.get_or_add_rPr()
    for tag, name in (("a:latin", latin), ("a:ea", ea), ("a:cs", latin)):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)


def tb(
    slide,
    l,
    t,
    w,
    h,
    text,
    *,
    size=14,
    bold=False,
    color=INK,
    serif=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    italic=False,
    wrap=True,
    ml=0.0,
    mt=0.0,
    mr=0.0,
    mb=0.0,
):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    tf.anchor = anchor
    tf.margin_left = Inches(ml)
    tf.margin_right = Inches(mr)
    tf.margin_top = Inches(mt)
    tf.margin_bottom = Inches(mb)
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(2 if i < len(lines) - 1 else 0)
        run = p.add_run()
        run.text = line
        apply_font(run, SERIF_L if serif else SANS_L, SERIF_E if serif else SANS_E, size, bold, color, italic)
    return box


def tb_runs(slide, l, t, w, h, runs, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, ml=0.0, mt=0.0):
    """runs: list of dicts with text, size, bold, color, serif, italic"""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.anchor = anchor
    tf.margin_left = Inches(ml)
    tf.margin_top = Inches(mt)
    tf.margin_right = Inches(0)
    tf.margin_bottom = Inches(0)
    first = True
    for spec in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = spec.get("align", align)
        p.space_before = Pt(spec.get("sb", 0))
        p.space_after = Pt(spec.get("sa", 0))
        p.line_spacing = spec.get("lh", 1.08)
        run = p.add_run()
        run.text = spec["text"]
        apply_font(
            run,
            SERIF_L if spec.get("serif") else SANS_L,
            SERIF_E if spec.get("serif") else SANS_E,
            spec.get("size", 14),
            spec.get("bold", False),
            spec.get("color", INK),
            spec.get("italic", False),
        )
    return box


def blank(prs, dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333333, 7.5, INK if dark else PAPER)
    return slide


def folio_light(slide, section=""):
    rect(slide, M, 0.42, 0.11, 0.28, CRIMSON)
    tb(slide, 0.74, 0.38, 6.4, 0.36, "每日时政科技简报", size=16, bold=True, color=INK, serif=True, anchor=MSO_ANCHOR.MIDDLE)
    right = FOLIO if not section else f"{FOLIO}  ·  {section}"
    tb(slide, 7.2, 0.38, 5.58, 0.36, right, size=11, color=SLATE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    rule(slide, M, 0.78, 13.333333 - 2 * M, INK, 0.014)
    rule(slide, M, 0.81, 13.333333 - 2 * M, CRIMSON, 0.006)


def folio_dark(slide, section=""):
    rect(slide, M, 0.42, 0.11, 0.28, CRIMSON)
    tb(slide, 0.74, 0.38, 6.4, 0.36, "每日时政科技简报", size=16, bold=True, color=PAPER, serif=True, anchor=MSO_ANCHOR.MIDDLE)
    right = FOLIO if not section else f"{FOLIO}  ·  {section}"
    tb(slide, 7.2, 0.38, 5.58, 0.36, right, size=11, color=RGBColor(0xB8, 0xB0, 0xA4), align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    rule(slide, M, 0.78, 13.333333 - 2 * M, PAPER, 0.01)


def why_line(text):
    return f"为何重要  {text}"


def card(slide, l, t, w, h, story, *, dark=False):
    bg = INK if dark else CARD
    ln = RGBColor(0x3A, 0x33, 0x2E) if dark else RULE
    fg = PAPER if dark else INK
    sub = RGBColor(0xC4, 0xBE, 0xB4) if dark else SLATE
    rect(slide, l, t, w, h, bg, ln, 0.75)
    pad = 0.22
    inner_l = l + pad
    inner_w = w - pad * 2
    rect(slide, inner_l, t + 0.18, 0.11, 0.22, CRIMSON)
    tb(slide, inner_l + 0.22, t + 0.14, inner_w - 0.22, 0.28, f"{story['no']}  ·  {story['kicker']}  ·  {story['source']}", size=11, color=CRIMSON, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(slide, inner_l, t + 0.48, inner_w, 0.92, story["title"], size=18, bold=True, color=fg, serif=True)
    tb(slide, inner_l, t + 1.44, inner_w, 1.55, story["summary"], size=14, color=fg)
    tb(slide, inner_l, t + 3.08, inner_w, 0.32, story.get("url", ""), size=10, color=sub)
    why_bg = RGBColor(0x24, 0x22, 0x20) if dark else RGBColor(0xE4, 0xDD, 0xD2)
    rect(slide, l + 0.12, t + h - 1.28, w - 0.24, 1.10, why_bg)
    tb(slide, inner_l, t + h - 1.18, inner_w, 0.24, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(slide, inner_l, t + h - 0.90, inner_w, 0.62, story["why"], size=13, color=fg)


def cover(prs):
    s = blank(prs, dark=True)
    folio_dark(s)
    tb(s, M, 1.05, 8.4, 0.32, "ASIA / SHANGHAI   MORNING EDITION", size=12, color=CRIMSON, bold=True)
    tb(s, M, 1.32, 9.2, 1.05, "每日时政科技简报", size=44, bold=True, color=PAPER, serif=True)
    tb(s, M, 2.28, 8.8, 0.42, "时政与科技的晨报 · 只收昨日首发", size=18, color=RGBColor(0xC9, 0xC2, 0xB6), serif=True)

    # giant date block
    rect(s, M, 2.92, 4.35, 3.55, RGBColor(0x24, 0x22, 0x20))
    tb(s, M + 0.22, 3.05, 3.9, 0.32, WEEKDAY, size=14, color=CRIMSON, bold=True)
    tb(s, M + 0.18, 3.28, 4.0, 1.7, "11.12", size=72, bold=True, color=PAPER, serif=True)
    tb(s, M + 0.22, 5.05, 3.9, 0.4, DATE_CN, size=16, color=PAPER, serif=True)
    tb(s, M + 0.22, 5.48, 3.9, 0.7, f"{ISSUE}    时政六则 · 科技五则", size=13, color=RGBColor(0xB8, 0xB0, 0xA4))

    # lead stack
    leads = [
        ("01", POLITICS[0]["title"]),
        ("07", TECH[0]["title"]),
        ("03", POLITICS[2]["title"]),
        ("08", TECH[1]["title"]),
    ]
    y = 2.92
    for no, title in leads:
        rect(s, 5.15, y, 7.63, 0.78, RGBColor(0x24, 0x22, 0x20))
        tb(s, 5.32, y + 0.08, 0.7, 0.62, no, size=22, bold=True, color=CRIMSON, serif=True, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 6.10, y + 0.12, 6.45, 0.56, title, size=16, color=PAPER, serif=True, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.90

    tb(s, M, 6.95, 12.2, 0.28, "编辑说明：新闻均核验为 2026-11-12 当日发生或当日首发，未收入更早稿件及 11月13日凌晨之后消息。", size=11, color=RGBColor(0x8A, 0x84, 0x7A))


def politics_opener(prs):
    s = blank(prs)
    folio_light(s, "时政")
    tb(s, M, 1.05, 2.2, 0.9, "01", size=72, bold=True, color=CRIMSON, serif=True)
    tb(s, 2.4, 1.22, 6.5, 0.7, "时政", size=48, bold=True, color=INK, serif=True)
    tb(s, M, 2.15, 7.4, 1.15, "昨日主轴是一场纪念大会、一场例行记者会，以及中美在经贸与政党两个层面同时展开的窗口期沟通。国内则有成品油按机制下调。", size=16, color=SLATE)

    rect(s, M, 3.50, 7.45, 2.85, CARD, RULE)
    tb(s, 0.78, 3.68, 6.95, 0.28, "开篇引语", size=12, color=CRIMSON, bold=True)
    tb(s, 0.78, 4.02, 6.95, 1.55, f"“{POLITICS[0]['quote']}”", size=20, color=INK, serif=True)
    tb(s, 0.78, 5.70, 6.95, 0.36, f"— {POLITICS[0]['quote_from']}  ·  纪念孙中山先生诞辰160周年大会", size=13, color=SLATE)

    # right index
    rect(s, 8.30, 1.10, 4.48, 5.25, INK)
    tb(s, 8.52, 1.28, 4.1, 0.32, "本日时政目录", size=13, color=CRIMSON, bold=True)
    y = 1.72
    for st in POLITICS:
        tb(s, 8.52, y, 0.55, 0.62, st["no"], size=18, bold=True, color=CRIMSON, serif=True)
        tb(s, 9.10, y, 3.42, 0.70, st["title"], size=13, color=PAPER, serif=True)
        y += 0.72
    tb(s, 8.52, 5.90, 4.05, 0.28, "六则  ·  经交叉核验", size=12, color=RGBColor(0xB8, 0xB0, 0xA4))


def lead_sun(prs):
    s = blank(prs)
    folio_light(s, "时政 · 01")
    st = POLITICS[0]
    tb(s, M, 1.00, 8.6, 0.28, f"{st['kicker']}  ·  {st['source']}", size=13, color=CRIMSON, bold=True)
    tb(s, M, 1.28, 8.4, 0.95, "纪念孙中山先生诞辰160周年\n大会在京举行", size=28, bold=True, color=INK, serif=True)
    tb(s, M, 2.32, 8.4, 1.25, st["summary"], size=16, color=INK)
    rect(s, M, 3.70, 8.4, 1.55, CARD, RULE)
    tb(s, 0.75, 3.84, 8.0, 0.24, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 0.75, 4.12, 8.0, 0.55, st["why"], size=15, color=INK)
    tb(s, 0.75, 4.72, 8.0, 0.38, st["url"], size=11, color=MUTED)

    rect(s, 9.20, 1.05, 3.58, 5.30, INK)
    tb(s, 9.40, 1.25, 3.2, 0.7, "160", size=56, bold=True, color=CRIMSON, serif=True)
    tb(s, 9.40, 1.95, 3.2, 0.4, "周年", size=20, color=PAPER, serif=True)
    tb(s, 9.40, 2.50, 3.2, 1.6, "1866 – 2026\n人民大会堂\n约 3000 人出席", size=16, color=PAPER)
    rule(s, 9.40, 4.25, 3.18, CRIMSON, 0.01)
    tb(s, 9.40, 4.45, 3.2, 1.55, "习近平强调学习继承发扬先生宝贵精神，把强国建设、民族复兴伟业不断推向前进。", size=14, color=RGBColor(0xC9, 0xC2, 0xB6))


def diplomacy_cards(prs):
    s = blank(prs)
    folio_light(s, "时政 · 02 / 06")
    tb(s, M, 0.95, 12.2, 0.42, "外交往来", size=36, bold=True, color=INK, serif=True)
    card(s, M, 1.55, 6.00, 5.25, POLITICS[1])
    card(s, 7.00, 1.55, 5.78, 5.25, POLITICS[5])


def two_col_us(prs):
    s = blank(prs)
    folio_light(s, "时政 · 03 / 04")
    tb(s, M, 0.95, 12.2, 0.42, "中美窗口", size=36, bold=True, color=INK, serif=True)
    # left
    st = POLITICS[2]
    rect(s, M, 1.55, 5.95, 5.25, CARD, RULE)
    tb(s, 0.78, 1.75, 5.5, 0.28, f"{st['no']}  ·  {st['kicker']}", size=12, color=CRIMSON, bold=True)
    tb(s, 0.78, 2.08, 5.5, 1.05, st["title"], size=22, bold=True, color=INK, serif=True)
    tb(s, 0.78, 3.20, 5.5, 1.55, st["summary"], size=15, color=INK)
    rect(s, 0.78, 4.90, 5.5, 1.55, INK)
    tb(s, 0.96, 5.05, 5.15, 0.3, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 0.96, 5.38, 5.15, 0.85, st["why"], size=14, color=PAPER)
    # right
    st = POLITICS[3]
    rect(s, 6.85, 1.55, 5.93, 5.25, INK)
    tb(s, 7.08, 1.75, 5.5, 0.28, f"{st['no']}  ·  {st['kicker']}", size=12, color=CRIMSON, bold=True)
    tb(s, 7.08, 2.08, 5.5, 1.05, st["title"], size=22, bold=True, color=PAPER, serif=True)
    tb(s, 7.08, 3.20, 5.5, 1.55, st["summary"], size=15, color=PAPER)
    tb(s, 7.08, 4.95, 5.5, 0.3, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 7.08, 5.28, 5.5, 1.15, st["why"], size=14, color=RGBColor(0xC9, 0xC2, 0xB6))


def oil_page(prs):
    s = blank(prs)
    folio_light(s, "时政 · 05")
    st = POLITICS[4]
    tb(s, M, 0.95, 8.5, 0.28, f"{st['kicker']}  ·  {st['source']}", size=13, color=CRIMSON, bold=True)
    tb(s, M, 1.28, 8.2, 1.15, st["title"], size=32, bold=True, color=INK, serif=True)
    tb(s, M, 2.55, 8.2, 1.35, st["summary"], size=16, color=INK)
    tb(s, M, 4.00, 8.2, 0.7, why_line(st["why"]), size=14, color=SLATE, italic=True)

    # number tiles
    tiles = [("-645", "元/吨", "汽油标准品"), ("-625", "元/吨", "柴油标准品"), ("0.51", "元/升", "92号汽油约降"), ("26", "元", "加满一箱约少花")]
    x = M
    for num, unit, lab in tiles:
        rect(s, x, 4.85, 2.90, 1.95, CARD, RULE)
        tb(s, x + 0.15, 5.00, 2.6, 0.7, num, size=32, bold=True, color=CRIMSON, serif=True)
        tb(s, x + 0.15, 5.68, 2.6, 0.35, unit, size=14, color=INK)
        tb(s, x + 0.15, 6.08, 2.6, 0.5, lab, size=13, color=SLATE)
        x += 3.15


def tech_opener(prs):
    s = blank(prs)
    folio_light(s, "科技")
    tb(s, M, 1.05, 2.2, 0.9, "02", size=72, bold=True, color=CRIMSON, serif=True)
    tb(s, 2.4, 1.22, 6.5, 0.7, "科技", size=48, bold=True, color=INK, serif=True)
    tb(s, M, 2.15, 7.5, 1.2, "昨日主线是算力权力的重组：英伟达换帅，谷歌与AMD加推专用芯片，百度把增长叙事转向推理，科学侧则交出一份脑图谱。", size=16, color=SLATE)

    rect(s, M, 3.55, 7.45, 2.80, CARD, RULE)
    tb(s, 0.78, 3.75, 6.95, 0.28, "观察", size=12, color=CRIMSON, bold=True)
    tb(s, 0.78, 4.10, 6.95, 1.90, "训练军备仍在，但推理、智能体与机密计算正在改写芯片规格书。治理风险与产品发布叠在同一天，市场用股价投票，实验室则继续绘制更细的地图。", size=16, color=INK)

    rect(s, 8.30, 1.10, 4.48, 5.25, INK)
    tb(s, 8.52, 1.28, 4.1, 0.32, "本日科技目录", size=13, color=CRIMSON, bold=True)
    y = 1.80
    for st in TECH:
        tb(s, 8.52, y, 0.55, 0.68, st["no"], size=18, bold=True, color=CRIMSON, serif=True)
        tb(s, 9.10, y, 3.42, 0.75, st["title"], size=13, color=PAPER, serif=True)
        y += 0.82
    tb(s, 8.52, 5.90, 4.05, 0.28, "五则  ·  经交叉核验", size=12, color=RGBColor(0xB8, 0xB0, 0xA4))


def nvidia_lead(prs):
    s = blank(prs)
    folio_light(s, "科技 · 07")
    st = TECH[0]
    tb(s, M, 0.98, 7.8, 0.28, f"{st['kicker']}  ·  {st['source']}", size=13, color=CRIMSON, bold=True)
    tb(s, M, 1.28, 8.0, 0.95, "黄仁勋辞去英伟达总裁、\n首席执行官及董事职务", size=26, bold=True, color=INK, serif=True)
    tb(s, M, 2.32, 8.0, 1.35, st["summary"], size=15, color=INK)
    rect(s, M, 3.78, 8.0, 1.70, CARD, RULE)
    tb(s, 0.75, 3.92, 7.6, 0.24, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 0.75, 4.20, 7.6, 0.55, st["why"], size=15, color=INK)
    tb(s, 0.75, 4.82, 7.6, 0.48, st["url"], size=11, color=MUTED)

    rect(s, 8.85, 1.05, 3.93, 5.30, INK)
    tb(s, 9.05, 1.28, 3.55, 0.3, "市场瞬时", size=12, color=CRIMSON, bold=True)
    tb(s, 9.05, 1.62, 3.55, 1.1, st["stat"], size=54, bold=True, color=PAPER, serif=True)
    tb(s, 9.05, 2.72, 3.55, 0.4, st["stat_label"], size=14, color=RGBColor(0xC9, 0xC2, 0xB6))
    rule(s, 9.05, 3.25, 3.5, CRIMSON, 0.01)
    tb(s, 9.05, 3.45, 3.55, 0.45, "收跌 4.3%", size=22, bold=True, color=PAPER, serif=True)
    tb(s, 9.05, 3.95, 3.55, 0.9, "当日市值蒸发逾1800亿美元。临时CEO克雷斯同时兼任CFO。", size=14, color=PAPER)
    tb(s, 9.05, 5.05, 3.55, 1.05, "11月17日仍按计划发布2027财年第三季报。", size=14, color=RGBColor(0xC9, 0xC2, 0xB6))


def chips_two_col(prs):
    s = blank(prs)
    folio_light(s, "科技 · 08 / 09")
    tb(s, M, 0.95, 12.2, 0.42, "芯片对垒", size=36, bold=True, color=INK, serif=True)
    st = TECH[1]
    rect(s, M, 1.55, 5.95, 5.25, CARD, RULE)
    tb(s, 0.78, 1.72, 5.5, 0.28, f"{st['no']}  ·  {st['kicker']}  ·  {st['source']}", size=12, color=CRIMSON, bold=True)
    tb(s, 0.78, 2.05, 5.5, 1.05, st["title"], size=20, bold=True, color=INK, serif=True)
    tb(s, 0.78, 3.15, 5.5, 1.7, st["summary"], size=14, color=INK)
    tb(s, 0.78, 4.95, 5.5, 0.28, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 0.78, 5.28, 5.5, 1.15, st["why"], size=14, color=SLATE)

    st = TECH[2]
    rect(s, 6.85, 1.55, 5.93, 5.25, INK)
    tb(s, 7.08, 1.72, 5.5, 0.28, f"{st['no']}  ·  {st['kicker']}  ·  {st['source']}", size=12, color=CRIMSON, bold=True)
    tb(s, 7.08, 2.05, 5.5, 1.05, st["title"], size=20, bold=True, color=PAPER, serif=True)
    tb(s, 7.08, 3.15, 5.5, 1.7, st["summary"], size=14, color=PAPER)
    tb(s, 7.08, 4.95, 5.5, 0.28, "为何重要", size=12, color=CRIMSON, bold=True)
    tb(s, 7.08, 5.28, 5.5, 1.15, st["why"], size=14, color=RGBColor(0xC9, 0xC2, 0xB6))


def baidu_nih(prs):
    s = blank(prs)
    folio_light(s, "科技 · 10 / 11")
    tb(s, M, 0.95, 12.2, 0.42, "产业与科学", size=36, bold=True, color=INK, serif=True)
    card(s, M, 1.55, 6.00, 5.25, TECH[3])
    card(s, 7.00, 1.55, 5.78, 5.25, TECH[4])


def closer(prs):
    s = blank(prs, dark=True)
    folio_dark(s, "收束")
    tb(s, M, 0.98, 12.2, 0.5, "收束 · 昨日十一则", size=36, bold=True, color=PAPER, serif=True)
    items = POLITICS + TECH
    # two columns of one-liners
    left, right = items[:6], items[6:]
    y = 1.65
    for st in left:
        tb(s, M, y, 0.5, 0.42, st["no"], size=14, bold=True, color=CRIMSON, serif=True)
        tb(s, 1.10, y, 5.4, 0.42, st["title"], size=14, color=PAPER, serif=True)
        y += 0.48
    y = 1.65
    for st in right:
        tb(s, 6.90, y, 0.5, 0.42, st["no"], size=14, bold=True, color=CRIMSON, serif=True)
        tb(s, 7.45, y, 5.3, 0.42, st["title"], size=14, color=PAPER, serif=True)
        y += 0.48

    rect(s, M, 5.55, 12.23, 1.45, RGBColor(0x24, 0x22, 0x20))
    tb(s, 0.78, 5.70, 11.7, 0.28, "编辑手记", size=12, color=CRIMSON, bold=True)
    tb(
        s,
        0.78,
        6.02,
        11.7,
        0.78,
        "本期按 Asia/Shanghai 收录 2026-11-12 当日发生或首发新闻，时政六则、科技五则。微软365全球故障发生于13日凌晨（上海时间），故未入选。AMD推理对比为其自行测试，已在稿中标明。",
        size=14,
        color=PAPER,
    )


def build(path: Path):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    cover(prs)
    politics_opener(prs)
    lead_sun(prs)
    diplomacy_cards(prs)
    two_col_us(prs)
    oil_page(prs)
    tech_opener(prs)
    nvidia_lead(prs)
    chips_two_col(prs)
    baidu_nih(prs)
    closer(prs)
    prs.save(str(path))
    print("wrote", path, "slides", len(prs.slides))


if __name__ == "__main__":
    out = Path(__file__).resolve().parent / "briefing.pptx"
    build(out)
