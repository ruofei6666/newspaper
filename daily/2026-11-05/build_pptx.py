#!/usr/bin/env python3
"""Generate daily/2026-11-05/briefing.pptx — 第045期晨报。"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

INK = RGBColor(0x1A, 0x1A, 0x1A)
PAPER = RGBColor(0xF7, 0xF4, 0xEF)
CRIMSON = RGBColor(0xC4, 0x1E, 0x3A)
SLATE = RGBColor(0x4A, 0x55, 0x68)
CARD = RGBColor(0xFF, 0xFC, 0xF7)
RULE = RGBColor(0xD6, 0xCF, 0xC4)
MUTE = RGBColor(0x8A, 0x84, 0x7A)
WARM = RGBColor(0xC9, 0xC2, 0xB6)
DARK_CARD = RGBColor(0x24, 0x24, 0x24)
CREAM = RGBColor(0xE8, 0xE2, 0xD8)

W = Inches(13.333)
H = Inches(7.5)

SERIF = "Cambria"
SANS = "Calibri"


def set_run(run, text, font, size, color, bold=False, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", font)
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = etree.SubElement(rPr, qn("a:cs"))
    cs.set("typeface", font)


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def stroke_shape(shape, color, pt=0.75):
    shape.line.color.rgb = color
    shape.line.width = Pt(pt)


def rect(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    fill_shape(sh, fill)
    if line:
        stroke_shape(sh, line)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def rounded(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    fill_shape(sh, fill)
    if line:
        stroke_shape(sh, line)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    try:
        sh.adjustments[0] = 0.06
    except Exception:
        pass
    return sh


def folio(slide, dark=False):
    ink = PAPER if dark else INK
    slate = WARM if dark else SLATE
    rect(slide, 0.5, 0.62, 12.333, 0.018, CRIMSON)
    rect(slide, 0.5, 0.65, 12.333, 0.01, ink if not dark else PAPER)
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.28), Inches(5.2), Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    run = p.add_run()
    set_run(run, "每日时政科技简报", SERIF, 13, ink, bold=True)
    box = slide.shapes.add_textbox(Inches(5.4), Inches(0.28), Inches(5.3), Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    set_run(run, "第045期  ·  2026年11月5日  ·  星期四", SANS, 11, slate)
    box = slide.shapes.add_textbox(Inches(11.0), Inches(0.28), Inches(1.83), Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    set_run(run, "晨刊", SANS, 12, CRIMSON, bold=True)


def textbox(slide, l, t, w, h, lines, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.anchor = anchor
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get("align", PP_ALIGN.LEFT)
        p.space_before = Pt(ln.get("before", 0))
        p.space_after = Pt(ln.get("after", 0))
        run = p.add_run()
        set_run(
            run,
            ln["text"],
            ln.get("font", SANS),
            ln.get("size", 14),
            ln.get("color", INK),
            ln.get("bold", False),
            ln.get("italic", False),
        )
    return box


def add_blank(prs, dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    fill_shape(bg, INK if dark else PAPER)
    bg.line.fill.background()
    folio(slide, dark=dark)
    return slide


def card_block(slide, l, t, w, h, meta, title, src, body, why):
    rounded(slide, l, t, w, h, CARD, RULE)
    textbox(slide, l + 0.16, t + 0.10, w - 0.32, 0.24, [
        {"text": meta, "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(slide, l + 0.16, t + 0.34, w - 0.32, 0.70, [
        {"text": title, "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(slide, l + 0.16, t + 1.06, w - 0.32, 0.24, [
        {"text": src, "font": SANS, "size": 12, "color": SLATE, "italic": True},
    ])
    textbox(slide, l + 0.16, t + 1.32, w - 0.32, max(0.6, h - 2.42), [
        {"text": body, "font": SANS, "size": 14, "color": INK},
    ])
    rect(slide, l + 0.16, t + h - 1.08, w - 0.32, 0.01, RULE)
    textbox(slide, l + 0.16, t + h - 1.02, w - 0.32, 0.22, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(slide, l + 0.16, t + h - 0.78, w - 0.32, 0.68, [
        {"text": why, "font": SANS, "size": 14, "color": SLATE},
    ])


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # ---------- 1 COVER ----------
    s = add_blank(prs, dark=True)
    textbox(s, 0.5, 0.86, 6.4, 0.28, [
        {"text": "●  THE MORNING DISPATCH", "font": SANS, "size": 11, "color": WARM, "bold": True},
    ])
    textbox(s, 7.4, 0.86, 5.43, 0.28, [
        {"text": "ASIA / SHANGHAI　　昨日新闻", "font": SANS, "size": 11, "color": WARM, "align": PP_ALIGN.RIGHT},
    ])
    rect(s, 0.5, 1.20, 12.333, 0.03, CRIMSON)
    rect(s, 0.5, 1.26, 12.333, 0.012, PAPER)
    textbox(s, 0.5, 1.40, 12.3, 0.32, [
        {"text": "第零肆伍期  ·  二零二六年十一月五日  ·  星期四", "font": SANS, "size": 14, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.5, 1.78, 12.3, 1.10, [
        {"text": "每日时政科技简报", "font": SERIF, "size": 48, "color": PAPER, "bold": True},
    ])
    textbox(s, 0.5, 2.90, 12.3, 0.85, [
        {"text": "只收录该日发生或该日首发的稿件。第九届进博会在上海开幕。乌镇把十八项开放人工智能倡议写成闭幕成果。百度交出未经后训练的基座权重，月之暗面按承诺放开 2.8 万亿参数。", "font": SANS, "size": 16, "color": WARM},
    ])
    metrics = [
        ("06", "时政", "进博 · 修法 · 反腐"),
        ("06", "科技", "开源 · 世界模型 · 光计算"),
        ("12", "合计", "周四沪时首发可核验"),
    ]
    for i, (n, lab, sub) in enumerate(metrics):
        x = 0.5 + i * 4.11
        rounded(s, x, 3.95, 3.91, 1.85, DARK_CARD)
        textbox(s, x + 0.22, 4.06, 3.5, 0.85, [
            {"text": n, "font": SERIF, "size": 48, "color": CRIMSON, "bold": True},
        ])
        textbox(s, x + 0.22, 4.92, 3.5, 0.30, [
            {"text": lab, "font": SANS, "size": 18, "color": PAPER, "bold": True},
        ])
        textbox(s, x + 0.22, 5.26, 3.5, 0.32, [
            {"text": sub, "font": SANS, "size": 12, "color": MUTE},
        ])
    textbox(s, 0.5, 6.05, 12.3, 0.85, [
        {"text": "今日主线：进博会重申贸易平衡四条路径。国务院令给基层减负。乌镇峰会收官。开源权重在北京与圣克拉拉同时落地。", "font": SANS, "size": 15, "color": PAPER},
    ])

    # ---------- 2 时政开篇 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 1.6, 0.78, [
        {"text": "01", "font": SERIF, "size": 56, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.1, 0.92, 10.7, 0.52, [
        {"text": "时政开篇", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 2.1, 1.42, 10.7, 0.42, [
        {"text": "上海的展馆在星期四打开闸门。北京同一天签发国务院令、公布隐性腐败案例、接待朝鲜劳动党代表团。", "font": SANS, "size": 14, "color": SLATE},
    ])
    rounded(s, 0.5, 1.95, 12.333, 0.78, CARD, RULE)
    textbox(s, 0.68, 2.02, 12.0, 0.20, [
        {"text": "刊头摘句 · 丁薛祥", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.68, 2.24, 12.0, 0.40, [
        {"text": "「大家关注贸易失衡问题很正常，关键是怎么实现贸易平衡。」", "font": SERIF, "size": 16, "color": INK, "bold": True},
    ])
    facts = [
        ("4142", "境外参展企业", "展览面积首次突破 44 万平方米，世界 500 强和行业龙头首次超过 300 家。"),
        ("14+4", "行政法规清理", "国务院令修改 14 部、废止 4 部，取消或调整基层专项规划编制要求。"),
        ("13.28亿", "基本医保参保", "截至 9 月已超 2024 年全年，局方称全年将连续两年正增长。"),
    ]
    for i, (n, lab, sub) in enumerate(facts):
        x = 0.5 + i * 4.11
        rounded(s, x, 2.90, 3.91, 1.55, CARD, RULE)
        textbox(s, x + 0.16, 3.00, 3.6, 0.55, [
            {"text": n, "font": SERIF, "size": 28, "color": CRIMSON, "bold": True},
        ])
        textbox(s, x + 0.16, 3.52, 3.6, 0.24, [
            {"text": lab, "font": SANS, "size": 13, "color": INK, "bold": True},
        ])
        textbox(s, x + 0.16, 3.76, 3.6, 0.58, [
            {"text": sub, "font": SANS, "size": 12, "color": SLATE},
        ])
    rounded(s, 0.5, 4.58, 12.333, 2.50, CARD, RULE)
    textbox(s, 0.68, 4.68, 1.4, 0.80, [
        {"text": "01", "font": SERIF, "size": 40, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.68, 5.48, 1.4, 0.55, [
        {"text": "开放 · 上海", "font": SANS, "size": 13, "color": SLATE, "bold": True},
    ])
    textbox(s, 2.20, 4.66, 10.3, 0.22, [
        {"text": "新华社 / 外交部 · 5 日开幕", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.20, 4.88, 10.3, 0.62, [
        {"text": "第九届进博会开幕", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 2.20, 5.50, 10.3, 0.72, [
        {"text": "丁薛祥提出共享市场、深化投资、创新赋能、增进互信四点看法。4142 家境外企业参展。虹桥论坛首次发布《出口中国能力发展报告》。", "font": SANS, "size": 14, "color": INK},
    ])
    textbox(s, 2.20, 6.24, 10.3, 0.20, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.20, 6.44, 10.3, 0.48, [
        {"text": "进博会是中国在经贸逆流中重申开放与贸易平衡路径的最大窗口。", "font": SANS, "size": 14, "color": SLATE},
    ])

    # ---------- 3 时政双栏：修法 / 两高 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 12.3, 0.58, [
        {"text": "一部减负令，六件隐性腐败案", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 0.5, 1.38, 12.3, 0.28, [
        {"text": "一栏把形式主义规划从行政法规里删掉，一栏把“商业机会”写成可认定的受贿。", "font": SANS, "size": 14, "color": SLATE},
    ])
    card_block(
        s, 0.5, 1.78, 6.05, 5.22,
        "02　立法 · 北京",
        "国务院令修改废止行政法规",
        "来源：中国政府网 / 新华社",
        "李强签署国务院令，公布《国务院关于修改和废止部分行政法规的决定》，自公布之日起施行。决定修改 14 部、废止 4 部，取消或调整基层专项规划编制要求，并与生态环境法典、粮食安全保障法等衔接。",
        "以修法方式把基层减负和机构改革写进行政法规，影响规划编制与执法依据。",
    )
    card_block(
        s, 6.78, 1.78, 6.05, 5.22,
        "03　法治 · 北京",
        "两高发布隐性腐败典型案例",
        "来源：最高人民法院 / 中新社",
        "最高人民法院、最高人民检察院 5 日发布 6 件典型案例，覆盖商业机会、战略配售股、挂名领薪等。其中明确：通过职务行为获得且确定能获利、不实质经营、不承担风险的“商业机会”，应以实际获利认定受贿数额。",
        "把隐性腐败的司法认定从个案经验写成可复用的裁判规则。",
    )

    # ---------- 4 时政双栏：王毅 / 日本 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 12.3, 0.58, [
        {"text": "一堂党际课，一份安保答问", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 0.5, 1.38, 12.3, 0.28, [
        {"text": "一栏是中朝关系在五中全会之后的跟进，一栏是东京拟改写的安保三文件。", "font": SANS, "size": 14, "color": SLATE},
    ])
    card_block(
        s, 0.5, 1.78, 6.05, 5.22,
        "04　外交 · 北京",
        "王毅会见朝鲜劳动党代表团",
        "来源：外交部",
        "11 月 5 日，王毅在京会见金成男率领的朝鲜劳动党代表团。王毅称愿落实两党两国最高领导人共识，并介绍二十届五中全会情况。金成男表示朝方愿遵循最高领导人意志发展朝中友好。",
        "习近平 6 月访朝后的党际跟进，关系到半岛局势与中朝关系落地节奏。",
    )
    card_block(
        s, 6.78, 1.78, 6.05, 5.22,
        "05　安全 · 北京",
        "外交部批日本安保三文件",
        "来源：新华社 / 外交部",
        "郭嘉昆 5 日回应日本拟于 12 月通过新版“安保三文件”、继续将中国定位为“最大战略挑战”、未来五年防卫开支提至 70 万亿至 80 万亿日元时说，日方以防卫为名行军事扩张之实，必须高度警惕。",
        "日本安保文件改写将牵动东亚军备与中日安全预期。",
    )

    # ---------- 5 医保 + 缺口 ----------
    s = add_blank(prs)
    rounded(s, 0.5, 0.88, 12.333, 3.55, CARD, RULE)
    textbox(s, 0.70, 1.00, 1.6, 0.85, [
        {"text": "06", "font": SERIF, "size": 44, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.70, 1.85, 1.6, 0.70, [
        {"text": "民生 · 北京", "font": SANS, "size": 13, "color": SLATE, "bold": True},
    ])
    textbox(s, 2.40, 1.02, 10.1, 0.22, [
        {"text": "中新社 / 澎湃新闻 · 5 日通报", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.40, 1.26, 10.1, 0.70, [
        {"text": "医保参保将连续两年正增长", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 2.40, 2.00, 10.1, 1.15, [
        {"text": "国家医保局 5 日通报，截至 9 月基本医保参保 13.28 亿人，全年将连续两年正增长。局方称“居民医保七连降”主因是剔除重复参保、居民转职工及人口总量下降。医保钱包 2.0 截至 4 日跨省共济 202.61 万人次、15.82 亿元。", "font": SANS, "size": 14, "color": INK},
    ])
    textbox(s, 2.40, 3.20, 10.1, 0.22, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.40, 3.42, 10.1, 0.70, [
        {"text": "参保数据口径之争关系到筹资可持续与公众对医保制度的信任。", "font": SANS, "size": 14, "color": SLATE},
    ])
    rounded(s, 0.5, 4.60, 6.05, 2.40, CARD, RULE)
    textbox(s, 0.68, 4.72, 5.7, 0.22, [
        {"text": "信息不足 · 时政栏", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.68, 4.96, 5.7, 0.55, [
        {"text": "4 日旧闻不补收", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 0.68, 5.54, 5.7, 1.28, [
        {"text": "韩正出席 COP31、李强会见 WIPO、澜湄外长会均为 4 日稿。美国中期选举两院易手已见前刊。朝鲜导弹试射发生在 4 日，不单列。", "font": SANS, "size": 14, "color": SLATE},
    ])
    rounded(s, 6.78, 4.60, 6.05, 2.40, INK)
    textbox(s, 6.96, 4.72, 5.7, 0.22, [
        {"text": "对照 · 不重写前刊", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 6.96, 5.06, 5.7, 1.70, [
        {"text": "进博会预告与乌镇开幕已见前刊。本期只收 5 日开幕讲话、虹桥双旗舰报告首发，以及乌镇闭幕发布会上的 18 条倡议。外交部 4 日例会不混入郭嘉昆 5 日答问。", "font": SANS, "size": 14, "color": PAPER},
    ])

    # ---------- 6 科技开篇 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 1.6, 0.78, [
        {"text": "02", "font": SERIF, "size": 56, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.1, 0.92, 10.7, 0.52, [
        {"text": "科技开篇", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 2.1, 1.42, 10.7, 0.50, [
        {"text": "没有国内火箭在这一天升空。折进星期四的，是乌镇把开源写成闭幕文件，两家中国实验室按承诺交出权重。", "font": SANS, "size": 14, "color": SLATE},
    ])
    stats = [
        ("18条", "开放倡议", "乌镇首次发布《携手构建全球开放人工智能行动倡议》。"),
        ("280B", "文心基座", "ERNIE-5.5-Flash-Base 总参数 2800 亿，每 token 激活 140 亿。"),
        ("2.8T", "Kimi 权重", "K3.1 Pro 按 10 月承诺开放，BF16 权重大约 5.6TB。"),
        ("16B", "世界模型", "Cosmos 3 四款开放检查点各约 160 亿可训练参数。"),
    ]
    for i, (n, lab, sub) in enumerate(stats):
        x = 0.5 + i * 3.16
        rounded(s, x, 2.05, 3.02, 1.70, CARD, RULE)
        textbox(s, x + 0.12, 2.14, 2.78, 0.50, [
            {"text": n, "font": SERIF, "size": 26, "color": CRIMSON, "bold": True},
        ])
        textbox(s, x + 0.12, 2.64, 2.78, 0.24, [
            {"text": lab, "font": SANS, "size": 13, "color": INK, "bold": True},
        ])
        textbox(s, x + 0.12, 2.90, 2.78, 0.70, [
            {"text": sub, "font": SANS, "size": 12, "color": SLATE},
        ])
    rounded(s, 0.5, 3.95, 12.333, 3.05, CARD, RULE)
    textbox(s, 0.68, 4.08, 1.5, 0.80, [
        {"text": "07", "font": SERIF, "size": 40, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.68, 4.88, 1.5, 0.60, [
        {"text": "治理 · 乌镇", "font": SANS, "size": 13, "color": SLATE, "bold": True},
    ])
    textbox(s, 2.30, 4.08, 10.2, 0.22, [
        {"text": "新华社 / 浙江在线 · 5 日闭幕", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.30, 4.30, 10.2, 0.62, [
        {"text": "乌镇发布开放人工智能倡议", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 2.30, 4.96, 10.2, 0.88, [
        {"text": "2026 年世界互联网大会乌镇峰会 5 日闭幕。来自 138 个国家和地区约 1800 名嘉宾参会。大会首次发布《携手构建全球开放人工智能行动倡议》，提出 18 条行动。数字经济产业合作大会签约 116 个项目、金额超 1163 亿元。", "font": SANS, "size": 14, "color": INK},
    ])
    textbox(s, 2.30, 5.88, 10.2, 0.20, [
        {"text": "为何重要", "font": SANS, "size": 11, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 2.30, 6.10, 10.2, 0.55, [
        {"text": "乌镇把“开源开放”写成官方公共产品，将影响全球南方的人工智能治理话语。", "font": SANS, "size": 14, "color": SLATE},
    ])

    # ---------- 7 科技双栏：百度 / Kimi ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 12.3, 0.58, [
        {"text": "一边交出基座，一边放开万亿参数", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 0.5, 1.38, 12.3, 0.28, [
        {"text": "百度放下未经后训练的起点；月之暗面按承诺把 2.8 万亿参数放到可下载的位置。", "font": SANS, "size": 14, "color": SLATE},
    ])
    card_block(
        s, 0.5, 1.78, 6.05, 5.22,
        "08　开源 · 北京",
        "百度开源文心基座权重",
        "来源：百度文心 / Hugging Face",
        "百度 5 日发布 ERNIE-5.5-Flash-Base，开放预训练终点权重。模型为稀疏 MoE，总参数 2800 亿、每 token 激活 140 亿，在 33.5 万亿 token 上从零预训练，推理支持 100 万上下文。权重未经中训练与后训练。",
        "把未做后训练的基座权重公开，降低国内高效基座模型的复现与微调门槛。",
    )
    card_block(
        s, 6.78, 1.78, 6.05, 5.22,
        "09　开源 · 北京",
        "月之暗面开放 Kimi 权重",
        "来源：IT之家",
        "月之暗面 5 日按 10 月承诺开放 Kimi K3.1 Pro 权重。模型总参数 2.8 万亿、激活约 1040 亿，支持约 100 万 token 上下文。Agent Arena 开放权重模型排名第一；BF16 权重大约 5.6TB，本地部署门槛极高。",
        "超大规模开源权重改写智能体评测格局，但存储与算力仍把“可下载”和“可运行”分开。",
    )

    # ---------- 8 科技对照 + 据报 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.80, 12.3, 0.55, [
        {"text": "一辆仿真车，一束光纤", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    card_block(
        s, 0.5, 1.42, 6.05, 3.55,
        "10　仿真 · 圣克拉拉",
        "英伟达开源 Cosmos 3",
        "来源：Hugging Face / NVIDIA",
        "英伟达 5 日新增四款开放权重检查点，用于驾驶多传感器仿真与双臂机器人交互，各约 160 亿可训练参数，许可为 OpenMDW1.1。",
        "世界模型从演示走向可复用的开放检查点。",
    )
    card_block(
        s, 6.78, 1.42, 6.05, 3.55,
        "11　科学 · 东京",
        "光纤上的深度学习",
        "来源：东京大学 / Nature Computational Science",
        "中岛弘平团队与 NTT 用 200 米光纤回路演示 SS-DFA 与单映射架构，减少中间状态测量，尚未在最大语言模型上验证。",
        "若能芯片化，光计算可能降低训练与散热能耗。",
    )
    rounded(s, 0.5, 5.12, 7.95, 1.92, CARD, RULE)
    textbox(s, 0.66, 5.20, 7.6, 0.20, [
        {"text": "12　芯片 · 欧洲　据报", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.66, 5.40, 7.6, 0.55, [
        {"text": "据报寻租欧洲算力", "font": SERIF, "size": 36, "color": INK, "bold": True},
    ])
    textbox(s, 0.66, 5.96, 7.6, 0.48, [
        {"text": "CNBC 引述知情人士称，月之暗面讨论在欧洲租赁 Blackwell 集群，规模或达 8000 颗以上。公司未置评。", "font": SANS, "size": 14, "color": INK},
    ])
    textbox(s, 0.66, 6.46, 7.6, 0.48, [
        {"text": "为何重要　出口管制管不住海外云上的远程访问。", "font": SANS, "size": 14, "color": SLATE},
    ])
    rounded(s, 8.62, 5.12, 4.21, 1.92, CARD, RULE)
    textbox(s, 8.78, 5.22, 3.9, 0.20, [
        {"text": "信息不足", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 8.78, 5.46, 3.9, 1.45, [
        {"text": "无国内发射。高通 HBC、浙大《自然》、布朗多铁薄膜、斯坦福光子芯片、腾讯 Hy4 均为旧稿，不补收。", "font": SANS, "size": 14, "color": SLATE},
    ])

    # ---------- 9 收束 ----------
    s = add_blank(prs, dark=True)
    textbox(s, 0.5, 0.86, 6.4, 0.26, [
        {"text": "THE EVENING LEDGER", "font": SANS, "size": 11, "color": WARM, "bold": True},
    ])
    textbox(s, 8.2, 0.86, 4.63, 0.26, [
        {"text": "第045期 · 收束", "font": SANS, "size": 11, "color": WARM, "align": PP_ALIGN.RIGHT},
    ])
    rect(s, 0.5, 1.18, 12.333, 0.03, CRIMSON)
    rect(s, 0.5, 1.24, 12.333, 0.012, PAPER)
    textbox(s, 0.5, 1.38, 12.3, 0.70, [
        {"text": "闸门在上海打开，权重在北京放下", "font": SERIF, "size": 36, "color": PAPER, "bold": True},
    ])
    textbox(s, 0.5, 2.30, 12.3, 1.15, [
        {"text": "丁薛祥把贸易失衡写成四条可执行的平衡路径。国务院令删掉基层专项规划，两高把“商业机会”写成受贿。乌镇用 18 条倡议收官。百度交出未经后训练的基座，月之暗面放开 2.8 万亿参数。没有火箭；有一桩据报的欧洲算力，和一束还没做成芯片的光。", "font": SANS, "size": 14, "color": CREAM},
    ])
    rounded(s, 0.5, 3.60, 6.05, 2.35, DARK_CARD)
    textbox(s, 0.68, 3.70, 5.7, 0.24, [
        {"text": "时政六条", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 0.68, 3.98, 5.7, 1.85, [
        {"text": "1  第九届进博会开幕，丁薛祥谈贸易平衡\n2  国务院令修改 14 部、废止 4 部行政法规\n3  两高发布新型、隐性职务犯罪典型案例\n4  王毅会见朝鲜劳动党代表团\n5  外交部回应日本安保三文件\n6  医保局称基本医保参保将连续两年正增长", "font": SANS, "size": 13, "color": PAPER},
    ])
    rounded(s, 6.78, 3.60, 6.05, 2.35, DARK_CARD)
    textbox(s, 6.96, 3.70, 5.7, 0.24, [
        {"text": "科技六条 · 缺口", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])
    textbox(s, 6.96, 3.98, 5.7, 1.85, [
        {"text": "1  乌镇峰会闭幕，发布开放 AI 18 条倡议\n2  百度开源 ERNIE-5.5-Flash-Base\n3  月之暗面开放 Kimi K3.1 Pro 权重\n4  英伟达开源 Cosmos 3 世界模型\n5  东京大学光学物理深度学习框架\n6  据报月之暗面寻求欧洲 Blackwell 算力", "font": SANS, "size": 13, "color": PAPER},
    ])
    textbox(s, 0.5, 6.10, 12.3, 0.70, [
        {"text": "日期窗口：Asia/Shanghai 2026-11-05。交叉来源：新华社、外交部、中国政府网、最高人民法院、中新社、百度文心、IT之家、NVIDIA、Nature Computational Science、CNBC。排除 4 日 COP31、澜湄外长会、中期选举揭晓及 3 日旧闻。", "font": SANS, "size": 12, "color": WARM},
    ])
    textbox(s, 0.5, 6.78, 12.3, 0.28, [
        {"text": "每日时政科技简报 · 第045期 · 完", "font": SANS, "size": 12, "color": CRIMSON, "bold": True},
    ])

    out = "/agent/repos/newspaper/daily/2026-11-05/briefing.pptx"
    prs.save(out)
    print("wrote", out, "slides", len(prs.slides))


if __name__ == "__main__":
    build()
