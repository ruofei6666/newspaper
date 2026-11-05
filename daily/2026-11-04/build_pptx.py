#!/usr/bin/env python3
"""Generate daily/2026-11-04/briefing.pptx — 第044期晨报。"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
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
SLIDE_W = 13.333
SLIDE_H = 7.5

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


def add_textbox(slide, l, t, w, h, text, font, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    set_run(p.add_run() if p.runs else (p.runs[0] if False else p.add_run()), text, font, size, color, bold, italic)
    # first paragraph already has an empty run created by add_run on empty para — clean
    return box


def write_para(tf, text, font, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT, space_after=4, first=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(space_after)
    run = p.add_run()
    set_run(run, text, font, size, color, bold, italic)
    return p


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
    # less rounded
    try:
        sh.adjustments[0] = 0.06
    except Exception:
        pass
    return sh


def folio(slide, dark=False):
    ink = PAPER if dark else INK
    slate = WARM if dark else SLATE
    # thin crimson rule + ink hairline, not a color band
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
    set_run(run, "第044期  ·  2026年11月4日  ·  星期三", SANS, 11, slate)
    box = slide.shapes.add_textbox(Inches(11.0), Inches(0.28), Inches(1.83), Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    set_run(run, "晨刊", SANS, 12, CRIMSON, bold=True)


def textbox(slide, l, t, w, h, lines, anchor=MSO_ANCHOR.TOP):
    """lines: list of dicts {text, font, size, color, bold, italic, align, after}"""
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
    textbox(slide, l + 0.18, t + 0.12, w - 0.36, 0.26, [
        {"text": meta, "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(slide, l + 0.18, t + 0.38, w - 0.36, 1.40, [
        {"text": title, "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(slide, l + 0.18, t + 1.80, w - 0.36, 0.28, [
        {"text": src, "font": SANS, "size": 12, "color": SLATE, "italic": True, "after": 0},
    ])
    textbox(slide, l + 0.18, t + 2.10, w - 0.36, 1.35, [
        {"text": body, "font": SANS, "size": 14, "color": INK, "after": 0},
    ])
    rect(slide, l + 0.18, t + h - 1.18, w - 0.36, 0.012, RULE)
    textbox(slide, l + 0.18, t + h - 1.12, w - 0.36, 0.22, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(slide, l + 0.18, t + h - 0.88, w - 0.36, 0.76, [
        {"text": why, "font": SANS, "size": 14, "color": SLATE, "after": 0},
    ])


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # ---------- 1 COVER ----------
    s = add_blank(prs, dark=True)
    textbox(s, 0.5, 0.86, 6.4, 0.28, [
        {"text": "●  THE MORNING DISPATCH", "font": SANS, "size": 11, "color": WARM, "bold": True, "after": 0},
    ])
    textbox(s, 7.4, 0.86, 5.43, 0.28, [
        {"text": "ASIA / SHANGHAI　　昨日新闻", "font": SANS, "size": 11, "color": WARM, "align": PP_ALIGN.RIGHT, "after": 0},
    ])
    rect(s, 0.5, 1.20, 12.333, 0.03, CRIMSON)
    rect(s, 0.5, 1.26, 12.333, 0.012, PAPER)
    textbox(s, 0.5, 1.40, 12.3, 0.32, [
        {"text": "第零肆肆期  ·  二零二六年十一月四日  ·  星期三", "font": SANS, "size": 14, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 1.78, 12.3, 1.15, [
        {"text": "每日时政科技简报", "font": SERIF, "size": 48, "color": PAPER, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 2.95, 12.3, 0.85, [
        {"text": "只收录该日发生或该日首发的稿件。民主党把参众两院从特朗普手里夺回来。韩正将飞往安塔利亚，中方就中俄美元首会晤保持沟通。国行 iPhone 第一次把千问写进系统设置。", "font": SANS, "size": 16, "color": WARM, "after": 0},
    ])
    metrics = [
        ("06", "时政", "国会 · 峰会 · 再贷款"),
        ("06", "科技", "千问 · 算力 · 5G"),
        ("12", "合计", "周三沪时首发可核验"),
    ]
    for i, (n, lab, sub) in enumerate(metrics):
        x = 0.5 + i * 4.11
        rounded(s, x, 4.00, 3.91, 1.85, DARK_CARD)
        textbox(s, x + 0.22, 4.12, 3.5, 0.85, [
            {"text": n, "font": SERIF, "size": 48, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.22, 4.98, 3.5, 0.32, [
            {"text": lab, "font": SANS, "size": 18, "color": PAPER, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.22, 5.32, 3.5, 0.32, [
            {"text": sub, "font": SANS, "size": 12, "color": MUTE, "after": 0},
        ])
    textbox(s, 0.5, 6.10, 12.3, 0.85, [
        {"text": "今日主线：美国中期选举改写华盛顿。央行把回购增持再贷款利率压到 1%。苹果在华开测系统级智能。英伟达中国订单只交出约百分之一。", "font": SANS, "size": 15, "color": PAPER, "after": 0},
    ])

    # ---------- 2 时政开篇 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.85, 2.2, 0.85, [
        {"text": "01", "font": SERIF, "size": 64, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.7, 0.98, 10.1, 0.62, [
        {"text": "时政开篇", "font": SERIF, "size": 40, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 2.7, 1.62, 10.1, 0.55, [
        {"text": "投票站在美东关闭之后，计票结果折进上海的星期三。北京同一天有外交部例会、总理会见、政协闭幕和央行稳市令。广州的广交会在七十周年这一天收官。", "font": SANS, "size": 15, "color": SLATE, "after": 0},
    ])
    rounded(s, 0.5, 2.32, 12.333, 1.15, CARD, RULE)
    textbox(s, 0.72, 2.42, 11.9, 0.24, [
        {"text": "刊头摘句 · 外交部", "font": SANS, "size": 11, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.72, 2.70, 11.9, 0.62, [
        {"text": "「中方就中俄美元首会晤同相关方保持着沟通。」——发言人毛宁，2026年11月4日例行记者会", "font": SERIF, "size": 18, "color": INK, "bold": True, "after": 0},
    ])
    facts = [
        ("两院", "民主党夺回国会", "翻转得州等七个参院席位，并越过众院 218 席门槛。阿拉斯加、密西西比 4 日仍未定。"),
        ("1%", "再贷款利率", "央行自 4 日起将股票回购增持再贷款利率下调 0.25 个百分点，银行已申请额度超 2000 亿元。"),
        ("35.2万", "广交会采购商", "第 140 届广交会闭幕，境外采购商来自 219 个国家和地区，较上届增 12.2%。"),
    ]
    for i, (n, lab, body) in enumerate(facts):
        x = 0.5 + i * 4.11
        rounded(s, x, 3.68, 3.91, 3.20, CARD, RULE)
        textbox(s, x + 0.22, 3.84, 3.5, 0.72, [
            {"text": n, "font": SERIF, "size": 36, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.22, 4.60, 3.5, 0.36, [
            {"text": lab, "font": SANS, "size": 16, "color": SLATE, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.22, 5.05, 3.5, 1.55, [
            {"text": body, "font": SANS, "size": 14, "color": INK, "after": 0},
        ])

    # ---------- 3 选举头条 ----------
    s = add_blank(prs)
    rounded(s, 0.5, 0.90, 2.15, 6.10, CARD, RULE)
    textbox(s, 0.62, 1.15, 1.9, 1.1, [
        {"text": "01", "font": SERIF, "size": 54, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 2.30, 1.9, 0.9, [
        {"text": "选举\n华盛顿", "font": SANS, "size": 16, "color": SLATE, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 3.40, 1.9, 1.4, [
        {"text": "纽约时报\n美联社\n4 日揭晓", "font": SANS, "size": 13, "color": MUTE, "after": 0},
    ])
    textbox(s, 2.90, 0.95, 9.9, 0.28, [
        {"text": "纽约时报 / 美联社 · 4 日揭晓", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 1.28, 9.9, 1.35, [
        {"text": "民主党夺回美国国会两院，中期选举重挫特朗普", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 2.70, 9.9, 0.32, [
        {"text": "来源：纽约时报 / 美联社    cn.nytimes.com    politico.com", "font": SANS, "size": 12, "color": SLATE, "italic": True, "after": 0},
    ])
    textbox(s, 2.90, 3.10, 9.9, 1.55, [
        {"text": "当地时间 3 日的美国中期选举结果于 4 日揭晓。据美联社，民主党翻转得州、北卡、南卡、俄亥俄、堪萨斯、艾奥瓦、缅因等参院席位，并越过众院 218 席门槛，二十年来首次在共和党总统任内同时掌控两院。通胀、生活成本与伊朗战争被广泛视为主因。", "font": SANS, "size": 16, "color": INK, "after": 0},
    ])
    rect(s, 2.90, 4.80, 9.9, 0.014, RULE)
    textbox(s, 2.90, 4.95, 9.9, 0.28, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 5.28, 9.9, 0.85, [
        {"text": "华盛顿权力结构重塑，将直接影响对华经贸、援乌与中东政策，并压缩特朗普余下任期的立法空间。", "font": SANS, "size": 15, "color": SLATE, "after": 0},
    ])
    # three mini numbers
    bits = [("7", "参院翻转州"), ("218", "众院多数门槛"), ("20年", "首次两院制衡")]
    for i, (n, lab) in enumerate(bits):
        x = 2.90 + i * 3.35
        rounded(s, x, 6.20, 3.15, 0.80, CARD, RULE)
        textbox(s, x + 0.14, 6.28, 1.1, 0.62, [
            {"text": n, "font": SERIF, "size": 22, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 1.20, 6.40, 1.8, 0.45, [
            {"text": lab, "font": SANS, "size": 13, "color": SLATE, "bold": True, "after": 0},
        ])

    # ---------- 4 双栏 外交 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.85, 12.3, 0.55, [
        {"text": "一位副主席启程，一位总干事抵京", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 1.40, 12.3, 0.40, [
        {"text": "一栏是气候峰会与大国元首会晤的预告，一栏是知识产权规则被重新写入人工智能时代。", "font": SANS, "size": 14, "color": SLATE, "after": 0},
    ])
    card_block(
        s, 0.5, 1.90, 6.05, 5.10,
        "02　外交 · 北京",
        "韩正将出席 COP31，沟通三边会晤",
        "来源：外交部    mfa.gov.cn",
        "发言人毛宁 4 日宣布，韩正将作为习近平主席特别代表于 11 月 6 日至 12 日赴土耳其出席 COP31 安塔利亚领导人峰会并访问格鲁吉亚。被问及深圳 APEC 期间中俄美元首会晤传闻，毛宁称中方正与相关方保持沟通。",
        "气候多边议程与大国元首外交同期推进，APEC 前夕的三边会晤传闻牵动乌克兰与全球秩序预期。",
    )
    card_block(
        s, 6.78, 1.90, 6.05, 5.10,
        "03　治理 · 北京",
        "李强会见世界知识产权组织总干事邓鸿森",
        "来源：新华社    news.cn",
        "4 日下午，国务院总理李强在人民大会堂会见世界知识产权组织总干事邓鸿森。李强表示中方愿在全球知识产权治理中发挥更大作用，并就人工智能、前沿技术议题讨论。邓鸿森称中国已跻身全球创新指数前十。",
        "知识产权规则正被重新写入 AI 时代，中方与 WIPO 的议程将影响全球南方能力建设与技术治理。",
    )

    # ---------- 5 双栏 金融/政协 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.85, 12.3, 0.55, [
        {"text": "稳市令落地，常委会闭幕", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 1.40, 12.3, 0.40, [
        {"text": "一栏是央行把回购增持再贷款写成 1% 的资金成本，一栏是人民政协把五中全会精神写成明年工作主线。", "font": SANS, "size": 14, "color": SLATE, "after": 0},
    ])
    card_block(
        s, 0.5, 1.90, 6.05, 5.10,
        "04　金融 · 北京",
        "回购增持再贷款利率下调至 1%",
        "来源：中国人民银行    pbc.gov.cn",
        "央行 4 日宣布，自即日起将股票回购增持再贷款利率下调 0.25 个百分点至 1%，支持对象扩至 21 家全国性银行和 5 家系统重要性城商行，支持比例 100%。座谈会后银行已申请额度超 2000 亿元。",
        "稳市工具加力常态化，直接关系到 A 股流动性与上市公司市值管理成本。",
    )
    card_block(
        s, 6.78, 1.90, 6.05, 5.10,
        "05　政协 · 北京",
        "全国政协常委会第十九次会议闭幕",
        "来源：新华社    news.cn",
        "政协第十四届全国委员会常务委员会第十九次会议 4 日下午在京闭幕。王沪宁主持并讲话，要求把学习贯彻二十届五中全会精神作为重大政治任务，并把迎接中共二十一大作为 2027 年工作主线。",
        "全会精神进入政协系统落地阶段，并划定明年全国政协工作主轴。",
    )

    # ---------- 6 广交会 ----------
    s = add_blank(prs)
    rounded(s, 0.5, 0.90, 2.15, 6.10, CARD, RULE)
    textbox(s, 0.62, 1.15, 1.9, 1.1, [
        {"text": "06", "font": SERIF, "size": 54, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 2.30, 1.9, 0.9, [
        {"text": "外贸\n广州", "font": SANS, "size": 16, "color": SLATE, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 3.40, 1.9, 1.4, [
        {"text": "央视新闻\n人民日报\n4 日闭幕", "font": SANS, "size": 13, "color": MUTE, "after": 0},
    ])
    textbox(s, 2.90, 0.95, 9.9, 0.28, [
        {"text": "央视新闻 · 4 日闭幕", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 1.28, 9.9, 1.15, [
        {"text": "第 140 届广交会闭幕，境外采购商创纪录", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 2.50, 9.9, 0.30, [
        {"text": "来源：央视新闻 / 人民日报    tv.cctv.com", "font": SANS, "size": 12, "color": SLATE, "italic": True, "after": 0},
    ])
    textbox(s, 2.90, 2.88, 9.9, 1.25, [
        {"text": "创办 70 周年的第 140 届广交会 4 日在广州闭幕。展会分三期举行，吸引 219 个国家和地区 35.2 万名境外采购商，较上届增 12.2%。3.2 万家企业参展，新产品、绿色产品占比分别达 23.2% 和 26%。", "font": SANS, "size": 16, "color": INK, "after": 0},
    ])
    nums = [("70", "周年"), ("35.2万", "境外采购商"), ("12.2%", "较上届增长"), ("26%", "绿色产品")]
    for i, (n, lab) in enumerate(nums):
        x = 2.90 + i * 2.50
        rounded(s, x, 4.28, 2.35, 1.35, CARD, RULE)
        textbox(s, x + 0.12, 4.38, 2.1, 0.62, [
            {"text": n, "font": SERIF, "size": 26, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.12, 5.08, 2.1, 0.38, [
            {"text": lab, "font": SANS, "size": 13, "color": SLATE, "bold": True, "after": 0},
        ])
    textbox(s, 2.90, 5.80, 9.9, 0.28, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 6.10, 9.9, 0.80, [
        {"text": "广交会是观察中国外贸结构与新质生产力出海的窗口，采购商规模创纪录显示外需并未退潮。", "font": SANS, "size": 15, "color": SLATE, "after": 0},
    ])

    # ---------- 7 科技开篇 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.85, 2.2, 0.85, [
        {"text": "02", "font": SERIF, "size": 64, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.7, 0.98, 10.1, 0.62, [
        {"text": "科技开篇", "font": SERIF, "size": 40, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 2.7, 1.62, 10.1, 0.70, [
        {"text": "没有国内火箭在这一天升空。折进星期三的，是系统级大模型第一次写进国行 iPhone，一份 700 亿美元的芯片订单只交出约百分之一，以及一条从卫星宽带伸向地面基站的 5G 协议。", "font": SANS, "size": 15, "color": SLATE, "after": 0},
    ])
    stats = [
        ("12月", "正式版", "国行 Apple 智能计划随 iOS 27.3 正式面向中国用户推出。"),
        ("1%", "交付比例", "中国云厂订购约 600 万颗 RTX Pro 5500，未来四月仅能收到约 1%。"),
        ("97%", "数据中心", "AMD 第三季度数据中心营收同比近乎翻倍，至 85 亿美元。"),
        ("2027", "地面 5G", "SpaceX 与高通约定明年起部署数万座 Starlink Mobile 基站。"),
    ]
    for i, (n, lab, body) in enumerate(stats):
        x = 0.5 + i * 3.16
        rounded(s, x, 2.55, 3.00, 4.45, CARD, RULE)
        textbox(s, x + 0.16, 2.75, 2.68, 0.85, [
            {"text": n, "font": SERIF, "size": 36, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.16, 3.65, 2.68, 0.40, [
            {"text": lab, "font": SANS, "size": 16, "color": INK, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.16, 4.15, 2.68, 2.40, [
            {"text": body, "font": SANS, "size": 14, "color": SLATE, "after": 0},
        ])

    # ---------- 8 Apple ----------
    s = add_blank(prs)
    rounded(s, 0.5, 0.90, 2.15, 6.10, CARD, RULE)
    textbox(s, 0.62, 1.15, 1.9, 1.1, [
        {"text": "07", "font": SERIF, "size": 54, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 2.30, 1.9, 0.9, [
        {"text": "终端\n中国", "font": SANS, "size": 16, "color": SLATE, "bold": True, "after": 0},
    ])
    textbox(s, 0.62, 3.40, 1.9, 1.6, [
        {"text": "动点科技\nIT之家\n4 日开测", "font": SANS, "size": 13, "color": MUTE, "after": 0},
    ])
    textbox(s, 2.90, 0.95, 9.9, 0.28, [
        {"text": "动点科技 / IT之家 · 4 日开测", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 1.28, 9.9, 1.25, [
        {"text": "国行 Apple 智能接入千问，随 iOS 27.3 开测", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 2.62, 9.9, 0.30, [
        {"text": "来源：动点科技 / IT之家    technode.com    ithome.com", "font": SANS, "size": 12, "color": SLATE, "italic": True, "after": 0},
    ])
    textbox(s, 2.90, 3.00, 9.9, 1.45, [
        {"text": "苹果 4 日随 iOS 27.3 开发者预览版在中国大陆开放 Apple 智能测试。Siri、写作工具与视觉智能可调用阿里千问，实时翻译、照片消除等端侧功能同步上线。首批机型为 iPhone 15 Pro 及更新机型，正式版计划 12 月推出。", "font": SANS, "size": 16, "color": INK, "after": 0},
    ])
    tags = [("Siri", "接入千问问答"), ("写作工具", "改写校对生成"), ("视觉智能", "识物与提问"), ("12月", "正式版窗口")]
    for i, (n, lab) in enumerate(tags):
        x = 2.90 + i * 2.50
        rounded(s, x, 4.60, 2.35, 1.20, CARD, RULE)
        textbox(s, x + 0.12, 4.70, 2.1, 0.42, [
            {"text": n, "font": SERIF, "size": 18, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.12, 5.18, 2.1, 0.42, [
            {"text": lab, "font": SANS, "size": 13, "color": SLATE, "after": 0},
        ])
    textbox(s, 2.90, 5.98, 9.9, 0.24, [
        {"text": "为何重要", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 2.90, 6.24, 9.9, 0.70, [
        {"text": "苹果在华 AI 能力迟到两年后落地，系统级大模型接入将改写手机入口与数据合规格局。", "font": SANS, "size": 15, "color": SLATE, "after": 0},
    ])

    # ---------- 9 双栏 芯片 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.85, 12.3, 0.55, [
        {"text": "一边订不到货，一边把营收翻倍", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 1.40, 12.3, 0.40, [
        {"text": "中国云厂的合规芯片只交出约百分之一；太平洋对岸，AMD 把数据中心写成公司主业。", "font": SANS, "size": 14, "color": SLATE, "after": 0},
    ])
    card_block(
        s, 0.5, 1.90, 6.05, 5.10,
        "08　芯片 · 华盛顿 / 北京",
        "英伟达中国订单四月仅交付约 1%",
        "来源：彭博社 / 联合早报    zaobao.com.sg",
        "彭博社 4 日引述知情人士称，阿里、腾讯、字节等中国 AI 与云企业向英伟达订购约 600 万颗 RTX Pro 5500，订单总额约 700 亿美元。按已确认计划，未来四个月仅能收到总量约 1%。",
        "合规高端算力供给远低于需求，将迫使云厂加速转向国产推理芯片，并抬高训练成本。",
    )
    card_block(
        s, 6.78, 1.90, 6.05, 5.10,
        "09　芯片 · 圣克拉拉",
        "AMD 数据中心收入近乎翻倍",
        "来源：AMD 官网    ir.amd.com",
        "AMD 公布 2026 财年第三季度业绩：营收 134 亿美元，同比增 45%；数据中心营收 85 亿美元，同比增 97%，占营收 64%。增长由 EPYC 与 Instinct MI450 带动。公司预计第四季度营收约 170 亿美元。",
        "AI 算力需求正从英伟达单极转向双供应商格局，AMD 的放量将影响全球 GPU 定价与交付节奏。",
    )

    # ---------- 10 三栏 ----------
    s = add_blank(prs)
    textbox(s, 0.5, 0.82, 12.3, 0.50, [
        {"text": "一份蓝皮书，一座地面网，一块近存芯片", "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 1.32, 12.3, 0.36, [
        {"text": "乌镇给出官方口径的智能经济体量。SpaceX 把卫星宽带伸向地面基站。高通试图让 AI 加速器少依赖昂贵的 HBM。", "font": SANS, "size": 13, "color": SLATE, "after": 0},
    ])
    triples = [
        ("10　产业 · 乌镇",
         "乌镇发布 2026 蓝皮书",
         "来源：新华社 / 央视新闻",
         "两份蓝皮书 4 日在乌镇峰会发布。2025 年中国企业级 AI 智能体市场 212 亿元，具身智能约 9150 亿元；截至 6 月，924 款生成式 AI 服务完成备案。",
         "蓝皮书给出官方口径的产业体量，是观察中国智能经济发展阶段的基准文本。"),
        ("11　通信 · 圣迭戈",
         "SpaceX 高通共建地面 5G",
         "来源：高通 / CNBC",
         "SpaceX 选用 Dragonwing 平台建设 Starlink Mobile 全美地面 5G 网络。多年期协议覆盖数万基站，采用 64T64R Massive MIMO，计划 2027 年起部署。",
         "卫星宽带商进入地面蜂窝市场，将改写美国电信竞争与 6G 基础设施路线。"),
        ("12　芯片 · 圣迭戈",
         "高通 HBC 验证超 HBM4",
         "来源：TrendForce",
         "高通完成第一代 HBC 硬件验证，将 LPDDR 直接堆叠于计算芯片，称带宽每瓦效率超过 HBM4 五倍。该技术将用于 AI250，计划 2027 年上半年送样。",
         "若能量产，近存计算可能降低 AI 加速器对昂贵 HBM 的依赖。"),
    ]
    for i, (meta, title, src, body, why) in enumerate(triples):
        x = 0.5 + i * 4.18
        rounded(s, x, 1.78, 4.00, 5.22, CARD, RULE)
        textbox(s, x + 0.16, 1.90, 3.68, 0.28, [
            {"text": meta, "font": SANS, "size": 11, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.16, 2.22, 3.68, 1.15, [
            {"text": title, "font": SERIF, "size": 36, "color": INK, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.16, 3.42, 3.68, 0.28, [
            {"text": src, "font": SANS, "size": 12, "color": SLATE, "italic": True, "after": 0},
        ])
        textbox(s, x + 0.16, 3.72, 3.68, 1.60, [
            {"text": body, "font": SANS, "size": 14, "color": INK, "after": 0},
        ])
        rect(s, x + 0.16, 5.50, 3.68, 0.012, RULE)
        textbox(s, x + 0.16, 5.58, 3.68, 0.22, [
            {"text": "为何重要", "font": SANS, "size": 11, "color": CRIMSON, "bold": True, "after": 0},
        ])
        textbox(s, x + 0.16, 5.82, 3.68, 1.00, [
            {"text": why, "font": SANS, "size": 13, "color": SLATE, "after": 0},
        ])

    # ---------- 11 收束 ----------
    s = add_blank(prs, dark=True)
    textbox(s, 0.5, 0.86, 6.4, 0.28, [
        {"text": "THE EVENING LEDGER", "font": SANS, "size": 11, "color": WARM, "bold": True, "after": 0},
    ])
    textbox(s, 7.4, 0.86, 5.43, 0.28, [
        {"text": "第044期 · 收束", "font": SANS, "size": 11, "color": WARM, "align": PP_ALIGN.RIGHT, "after": 0},
    ])
    rect(s, 0.5, 1.18, 12.333, 0.03, CRIMSON)
    rect(s, 0.5, 1.24, 12.333, 0.012, PAPER)
    textbox(s, 0.5, 1.40, 12.3, 1.05, [
        {"text": "两院易手，再贷款压到百分之一，千问写进系统设置", "font": SERIF, "size": 36, "color": PAPER, "bold": True, "after": 0},
    ])
    textbox(s, 0.5, 2.50, 12.3, 1.15, [
        {"text": "民主党把参众两院从特朗普手里夺回来。韩正将飞往安塔利亚，中方承认正在沟通一场可能的中俄美元首会晤。央行把回购增持再贷款利率压到 1%。国行 iPhone 第一次把千问写进设置，600 万颗合规芯片只交出约百分之一。", "font": SANS, "size": 15, "color": CREAM, "after": 0},
    ])
    rounded(s, 0.5, 3.80, 6.05, 2.35, DARK_CARD)
    textbox(s, 0.70, 3.92, 5.7, 0.28, [
        {"text": "时政六条", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 0.70, 4.24, 5.7, 1.75, [
        {"text": "1  民主党夺回美国国会两院\n2  韩正将出席 COP31，中方沟通三边会晤\n3  李强会见 WIPO 总干事邓鸿森\n4  央行下调回购增持再贷款利率至 1%\n5  全国政协常委会第十九次会议闭幕\n6  第 140 届广交会闭幕，采购商创纪录", "font": SANS, "size": 13, "color": PAPER, "after": 0},
    ])
    rounded(s, 6.78, 3.80, 6.05, 2.35, DARK_CARD)
    textbox(s, 6.98, 3.92, 5.7, 0.28, [
        {"text": "科技六条", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])
    textbox(s, 6.98, 4.24, 5.7, 1.75, [
        {"text": "1  国行 Apple 智能接入千问开测\n2  英伟达中国订单四月内仅交付约 1%\n3  AMD 数据中心收入近乎翻倍\n4  世界互联网大会发布 2026 蓝皮书\n5  SpaceX 携高通建设地面 5G 网\n6  高通完成 HBC 近存计算验证", "font": SANS, "size": 13, "color": PAPER, "after": 0},
    ])
    textbox(s, 0.5, 6.28, 12.3, 0.55, [
        {"text": "日期窗口：Asia/Shanghai 2026-11-04。交叉来源：纽约时报、外交部、新华社、央行、央视、动点科技、彭博/早报、AMD、CNBC、TrendForce。排除 3 日外汇存准、DeepSeek V5 传闻、力箭二号发射及 5 日凌晨之后的追评。", "font": SANS, "size": 11, "color": WARM, "after": 0},
    ])
    textbox(s, 0.5, 6.88, 12.3, 0.28, [
        {"text": "每日时政科技简报 · 第044期 · 完", "font": SANS, "size": 12, "color": CRIMSON, "bold": True, "after": 0},
    ])

    out = "/agent/repos/newspaper/daily/2026-11-04/briefing.pptx"
    prs.save(out)
    print("wrote", out, "slides", len(prs.slides))


if __name__ == "__main__":
    build()
