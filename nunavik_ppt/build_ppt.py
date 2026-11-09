"""Build the NUNAVIK comprehensive development plan deck from the source PDF content."""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
ASSETS = HERE / "assets"
OUTPUT = HERE / "努纳维克NUNAVIK地区综合发展规划.pptx"

FONT = "Microsoft YaHei"

NAVY = RGBColor(0x0B, 0x1D, 0x33)
PANEL = RGBColor(0x14, 0x33, 0x55)
PANEL_DARK = RGBColor(0x0F, 0x26, 0x42)
GREEN = RGBColor(0x3D, 0xDC, 0x97)
TEAL = RGBColor(0x2E, 0xC4, 0xB6)
VIOLET = RGBColor(0x9B, 0x8A, 0xD6)
GOLD = RGBColor(0xF4, 0xC9, 0x5D)
ICE = RGBColor(0xEA, 0xF6, 0xFF)
MUTED = RGBColor(0xA9, 0xC1, 0xD6)
LINE = RGBColor(0x2A, 0x4E, 0x74)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
page_no = [0]


def set_fill(shape, color, alpha=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if alpha is not None:
        srgb = shape.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = etree.SubElement(srgb, qn("a:alpha"))
        a.set("val", str(int(alpha * 100000)))


def set_line(shape, color=None, width=1.0, alpha=None):
    if color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = color
    shape.line.width = Pt(width)
    if alpha is not None:
        srgb = shape.line._ln.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = etree.SubElement(srgb, qn("a:alpha"))
        a.set("val", str(int(alpha * 100000)))


def gradient(shape, c1, c2, angle=0):
    shape.fill.gradient()
    shape.fill.gradient_angle = angle
    stops = shape.fill.gradient_stops
    stops[0].color.rgb = c1
    stops[0].position = 0
    stops[1].color.rgb = c2
    stops[1].position = 1.0


def rect(slide, x, y, w, h, color=None, alpha=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE, radius=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if color is None:
        shp.fill.background()
    else:
        set_fill(shp, color, alpha)
    set_line(shp, line, line_w)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    shp.shadow.inherit = False
    shp.text_frame.text = ""
    return shp


def rrect(slide, x, y, w, h, color=PANEL, alpha=0.72, line=LINE, radius=0.06, line_w=1.0):
    return rect(slide, x, y, w, h, color, alpha, line, line_w, MSO_SHAPE.ROUNDED_RECTANGLE, radius)


def _apply_font(run, size, color, bold, font):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", font)


def text(slide, x, y, w, h, paras, size=16, color=ICE, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, font=FONT, spacing=1.15, space_after=0, shape=None):
    """paras: str | list[str | list[(text, {opts})]]"""
    if shape is None:
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        tb = shape
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(0.04))
    if isinstance(paras, str):
        paras = [paras]
    paras = [line for p in paras for line in (p.split("\n") if isinstance(p, str) else [p])]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        runs = [(para, {})] if isinstance(para, str) else para
        for txt, opt in runs:
            r = p.add_run()
            r.text = txt
            _apply_font(r, opt.get("size", size), opt.get("color", color), opt.get("bold", bold), opt.get("font", font))
    return tb


def arrow(slide, x1, y1, x2, y2, color=TEAL, width=2.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")
    return c


def chip(slide, x, y, w, h, label, color=TEAL, size=14, text_color=ICE, alpha=0.18, bold=False):
    s = rect(slide, x, y, w, h, color, alpha, color, 1.0, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
    text(slide, x, y, w, h, label, size=size, color=text_color, bold=bold, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=s)
    return s


def badge(slide, x, y, d, label, color=GREEN, size=16, text_color=NAVY):
    s = rect(slide, x, y, d, d, color, None, None, shape=MSO_SHAPE.OVAL)
    text(slide, x, y, d, d, label, size=size, color=text_color, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=s)
    s.text_frame.margin_left = s.text_frame.margin_right = 0
    return s


def background(slide, name):
    slide.shapes.add_picture(str(ASSETS / name), 0, 0, prs.slide_width, prs.slide_height)


def footer(slide):
    page_no[0] += 1
    text(slide, 0.7, 7.0, 6, 0.3, "努纳维克 NUNAVIK 地区综合发展规划", size=10, color=MUTED)
    text(slide, SW - 1.7, 7.0, 1.0, 0.3, f"{page_no[0]:02d}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)


def content_slide(tag, title, subtitle=None):
    s = prs.slides.add_slide(BLANK)
    background(s, "content.jpg")
    text(s, 0.7, 0.42, 8, 0.32, tag, size=12, color=TEAL, bold=True)
    text(s, 0.7, 0.72, 11.9, 0.7, title, size=28, color=ICE, bold=True)
    bar = rect(s, 0.72, 1.45, 1.3, 0.06, GREEN)
    gradient(bar, GREEN, VIOLET, 0)
    if subtitle:
        text(s, 2.2, 1.32, 10.4, 0.34, subtitle, size=13, color=MUTED)
    footer(s)
    return s


def section_slide(num, title, en, desc):
    s = prs.slides.add_slide(BLANK)
    background(s, "section.jpg")
    text(s, 0.9, 1.15, 4, 1.6, num, size=110, color=GOLD, bold=True, font="Arial")
    bar = rect(s, 1.0, 3.25, 1.6, 0.07, GREEN)
    gradient(bar, GREEN, VIOLET, 0)
    text(s, 0.9, 3.5, 8, 0.9, title, size=40, color=ICE, bold=True)
    text(s, 0.95, 4.4, 8, 0.4, en, size=14, color=TEAL, bold=True, font="Arial")
    text(s, 0.95, 4.95, 6.6, 1.0, desc, size=16, color=MUTED, spacing=1.35)
    page_no[0] += 1
    return s


# ---------------------------------------------------------------- 1. Cover
s = prs.slides.add_slide(BLANK)
background(s, "cover.jpg")
shade = rect(s, 0, 0, SW, 3.6, NAVY, 0.25)
text(s, 0.9, 0.75, 9, 0.4, "COMPREHENSIVE DEVELOPMENT PLAN", size=14, color=TEAL, bold=True, font="Arial")
text(s, 0.9, 1.15, 11, 1.1, "努纳维克 NUNAVIK", size=54, color=ICE, bold=True)
text(s, 0.9, 2.15, 11, 0.9, "地区综合发展规划", size=40, color=ICE, bold=True)
bar = rect(s, 0.95, 3.15, 2.0, 0.07, GREEN)
gradient(bar, GREEN, VIOLET, 0)
text(s, 0.9, 3.35, 11, 0.5, "教育建设  ·  绿色能源  ·  人才与资金  ·  极光旅游", size=18, color=ICE)
page_no[0] += 1

# ---------------------------------------------------------------- 2. Agenda
s = content_slide("CONTENTS", "目录 · 四大重点项目")
items = [
    ("01", "学校教学大楼建设", "引入中国建设力量\n成立筹备项目公司\n落实前期预算与资金", GREEN),
    ("02", "绿色能源项目", "风力发电 · 太阳能光伏\n小型水力发电\n建立储能公司", TEAL),
    ("03", "人才与移民基金", "向 UN 申请设立城市\n及移民基金计划\n引进人才与国际资金", VIOLET),
    ("04", "INUVIK 极光世界", "打造北美顶级旅游项目\n带动 GDP 与就业\n惠及当地居民", GOLD),
]
cw, gap, y0 = 2.8, 0.243, 2.05
for i, (num, title, desc, col) in enumerate(items):
    x = 0.7 + i * (cw + gap)
    rrect(s, x, y0, cw, 4.3)
    top = rect(s, x + 0.02, y0 + 0.02, cw - 0.04, 0.08, col)
    text(s, x + 0.3, y0 + 0.4, 2.2, 1.0, num, size=54, color=col, bold=True, font="Arial")
    text(s, x + 0.3, y0 + 1.55, cw - 0.4, 0.6, title, size=18, color=ICE, bold=True)
    rect(s, x + 0.32, y0 + 2.2, 0.5, 0.03, col)
    text(s, x + 0.3, y0 + 2.4, cw - 0.5, 1.6, desc, size=14, color=MUTED, spacing=1.45)

# ---------------------------------------------------------------- 3. Section 01
section_slide("01", "学校教学大楼建设项目", "SCHOOL BUILDING PROJECT",
              "借助中国建设力量，以项目公司为载体，\n系统推进高寒地区教学大楼的规划与建设。")

# ---------------------------------------------------------------- 4. Project 01
s = content_slide("PROJECT 01", "建设学校教学大楼的项目")
cw, gap, y0, ch = 3.77, 0.31, 1.95, 3.85
cols = [GREEN, TEAL, VIOLET]
heads = [("A", "是否需要使用中国的建设力量？"), ("B", "是否需要成立（筹备）项目公司？"), ("C", "前期工作预算与资金安排")]
for i, (k, h) in enumerate(heads):
    x = 0.7 + i * (cw + gap)
    rrect(s, x, y0, cw, ch)
    badge(s, x + 0.25, y0 + 0.25, 0.55, k, cols[i], 18)
    text(s, x + 0.95, y0 + 0.22, cw - 1.1, 0.62, h, size=16, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x + 0.25, y0 + 1.0, cw - 0.5, 0.015, LINE)

xA = 0.7
text(s, xA + 0.25, y0 + 1.15, cw - 0.5, 0.4, "覆盖建设全链条：", size=14, color=MUTED)
for j, lab in enumerate(["设计", "规划", "材料", "设备", "施工"]):
    cx = xA + 0.25 + (j % 3) * 1.1
    cy = y0 + 1.65 + (j // 3) * 0.62
    chip(s, cx, cy, 0.98, 0.46, lab, GREEN, 14)
text(s, xA + 0.25, y0 + 3.05, cw - 0.5, 0.7, "发挥中国在基建领域的综合优势", size=13, color=MUTED)

xB = 0.7 + cw + gap
text(s, xB + 0.25, y0 + 1.15, cw - 0.5, 0.7, "由项目公司负责前期准备：\n规划、初步设计", size=14, color=ICE, spacing=1.3)
text(s, xB + 0.25, y0 + 1.95, cw - 0.5, 0.35, "需赴现场了解：", size=13, color=MUTED)
for j, lab in enumerate(["冻土施工", "高寒地区保温", "极昼极夜条件下的采光"]):
    chip(s, xB + 0.25, y0 + 2.35 + j * 0.47, cw - 0.5, 0.38, lab, TEAL, 13)

xC = 0.7 + 2 * (cw + gap)
text(s, xC + 0.25, y0 + 1.15, cw - 0.5, 1.0, "前期工作需要编制预算，\n并做出资金安排", size=14, color=ICE, spacing=1.3)
box = rrect(s, xC + 0.25, y0 + 2.15, cw - 0.5, 1.35, VIOLET, 0.16, VIOLET, 0.1)
text(s, xC + 0.4, y0 + 2.2, cw - 0.8, 1.25,
     [[("资金安排纳入股权设计", {"bold": True, "color": ICE, "size": 15})],
      [("在成立项目公司（B）时，将前期预算设计在股权结构中", {"color": MUTED, "size": 12})]],
     anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)

steps = ["成立筹备项目公司", "现场调研", "规划与初步设计", "预算与资金安排", "引入中国建设力量"]
sy, sw_, sg = 6.12, 2.18, 0.25
for j, st in enumerate(steps):
    sx = 0.7 + j * (sw_ + sg)
    chip(s, sx, sy, sw_, 0.5, st, GREEN if j == 0 else TEAL, 13, alpha=0.12)
    if j < len(steps) - 1:
        arrow(s, sx + sw_ + 0.02, sy + 0.25, sx + sw_ + sg - 0.02, sy + 0.25, MUTED, 1.5)

# ---------------------------------------------------------------- 5. Section 02
section_slide("02", "绿色能源项目", "GREEN ENERGY PROJECT",
              "风、光、水多能互补，配套储能系统，\n为地区发展提供稳定、清洁的电力保障。")

# ---------------------------------------------------------------- 6. Project 02
s = content_slide("PROJECT 02", "绿色能源项目：多能互补 + 储能")
srcs = [("A", "风力发电", GREEN), ("B", "太阳能光伏发电", GOLD), ("C", "小型水力发电", TEAL)]
for j, (k, lab, col) in enumerate(srcs):
    y = 2.05 + j * 1.35
    rrect(s, 0.7, y, 2.6, 0.95, PANEL, 0.8, col, 0.12)
    badge(s, 0.88, y + 0.2, 0.55, k, col, 16)
    text(s, 1.5, y, 1.75, 0.95, lab, size=15, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 3.35, y + 0.475, 4.05, 3.82 + (j - 1) * 0.5, col, 2.0)
st = rrect(s, 4.1, 2.75, 2.2, 2.15, VIOLET, 0.22, VIOLET, 0.1, 1.5)
badge(s, 4.88, 2.95, 0.6, "D", VIOLET, 18)
text(s, 4.1, 3.62, 2.2, 1.2, [[("储能公司", {"bold": True, "size": 18})], [("平抑波动 · 储存电能", {"size": 12, "color": MUTED})]],
     align=PP_ALIGN.CENTER, spacing=1.3)
arrow(s, 6.35, 3.82, 6.95, 3.82, GREEN, 2.5)
out = rrect(s, 7.0, 2.95, 1.2, 1.75, GREEN, 0.2, GREEN, 0.12)
text(s, 7.0, 2.95, 1.2, 1.75, "稳定\n清洁\n电力", size=15, color=ICE, bold=True, align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)

rrect(s, 8.6, 1.95, 4.03, 4.75, PANEL_DARK, 0.75)
details = [
    ("A  风力发电", "机会更大一些；需要掌握风力强度、规律等数据", GREEN),
    ("B  太阳光伏发电", "作为清洁电源的重要组成部分", GOLD),
    ("C  小型水力发电", "因地制宜开发小型水电资源", TEAL),
    ("D  建立储能公司", "解决 A、B、C 中电力不稳及储能问题", VIOLET),
]
for j, (h, d, col) in enumerate(details):
    y = 2.15 + j * 1.12
    rect(s, 8.85, y + 0.06, 0.07, 0.85, col)
    text(s, 9.08, y, 3.4, 0.4, h, size=15, color=ICE, bold=True)
    text(s, 9.08, y + 0.42, 3.4, 0.6, d, size=12, color=MUTED, spacing=1.25)
star = rrect(s, 11.45, 2.17, 1.0, 0.32, GREEN, 0.25, GREEN, 0.5)
text(s, 11.45, 2.17, 1.0, 0.32, "优先方向", size=10, color=GREEN, bold=True, align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, shape=star)

# ---------------------------------------------------------------- 7. Section 03
section_slide("03", "人才与移民基金计划", "TALENT & IMMIGRATION FUND",
              "向 UN 申请设立“NWT”或因纽维克式城市、\n移民基金，引进人才与国际资金。")

# ---------------------------------------------------------------- 8. Project 03
s = content_slide("PROJECT 03", "向 UN 申请设立“NWT”或因纽维克式城市、移民基金")
rrect(s, 0.7, 1.95, 6.9, 4.75)
badge(s, 0.95, 2.18, 0.55, "A", GREEN, 18)
text(s, 1.65, 2.15, 5.7, 0.6, "引进人才：当地经济发展的关键", size=18, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.95, 2.9, 6.4, 0.35, "重点引进领域", size=13, color=MUTED)
fields = ["能源技术", "储能技术", "全球旅游规划设计", "城市管理", "教育及技术培训", "酒店管理", "物流", "金融"]
for j, f in enumerate(fields):
    cx = 0.95 + (j % 2) * 3.25
    cy = 3.35 + (j // 2) * 0.78
    chip(s, cx, cy, 3.05, 0.58, f, GREEN if j % 2 == 0 else TEAL, 15)

rrect(s, 7.9, 1.95, 4.73, 4.75)
badge(s, 8.15, 2.18, 0.55, "B", VIOLET, 18)
text(s, 8.85, 2.15, 3.7, 0.6, "引入国际资金", size=18, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
flow = [("通过移民", "从国际引入资金"), ("投入建设", "资金投入到 INUVIK 建设之中"), ("平衡缺口", "弥补地区发展的资金缺口")]
for j, (h, d) in enumerate(flow):
    y = 3.0 + j * 1.22
    rrect(s, 8.15, y, 4.23, 0.88, VIOLET, 0.14, VIOLET, 0.12)
    text(s, 8.35, y, 1.4, 0.88, h, size=15, color=GOLD, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 9.75, y, 2.55, 0.88, d, size=13, color=ICE, anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)
    if j < 2:
        arrow(s, 10.27, y + 0.9, 10.27, y + 1.2, VIOLET, 1.8)

# ---------------------------------------------------------------- 9. Section 04
section_slide("04", "“INUVIK 极光世界”旅游项目", "INUVIK AURORA WORLD",
              "建立北美（至少加拿大）最大、最豪华、\n最有吸引力的旅游项目。")

# ---------------------------------------------------------------- 10. Overview
s = content_slide("PROJECT 04 · 总览", "“INUVIK 极光世界”：五大主题板块",
                  "目标：北美（至少加拿大）最大、最豪华、最有吸引力的旅游项目")
zones = [
    ("A", "热带风情乐园", "不去夏威夷", "人造沙滩、温泉海浪\n场内四季如春", GREEN),
    ("B", "免税购物天堂", "不去拉斯维加斯", "豪华购物体验\n全球最齐备免税商品", GOLD),
    ("C", "环球美食街区", "不用周游世界", "极地海鲜、中国美食\n日料、泰餐……", TEAL),
    ("D", "极光晚会", "全球独一无二", "极光投射大屏\n音乐歌舞 + 中国烟花", VIOLET),
    ("E", "冰雪艺术世界", "世界冰雕名作", "长城、凯旋门、科隆教堂\n彩色冰雕世界", RGBColor(0x7F, 0xC8, 0xF8)),
]
cw, gap, y0 = 2.226, 0.2, 1.95
for i, (k, title, tagline, desc, col) in enumerate(zones):
    x = 0.7 + i * (cw + gap)
    rrect(s, x, y0, cw, 3.3)
    rect(s, x + 0.02, y0 + 0.02, cw - 0.04, 0.08, col)
    badge(s, x + cw / 2 - 0.33, y0 + 0.35, 0.66, k, col, 20)
    text(s, x + 0.1, y0 + 1.15, cw - 0.2, 0.45, title, size=17, color=ICE, bold=True, align=PP_ALIGN.CENTER)
    text(s, x + 0.1, y0 + 1.62, cw - 0.2, 0.35, tagline, size=12, color=col, bold=True, align=PP_ALIGN.CENTER)
    text(s, x + 0.12, y0 + 2.1, cw - 0.24, 1.1, desc, size=12, color=MUTED, align=PP_ALIGN.CENTER, spacing=1.35)
y1 = 5.5
benefits = [("F", "经济带动", "增加 GDP、就业与城市魅力"), ("G", "投资引入", "服务业、矿产、金融、新能源等"),
            ("H", "居民受益", "共享成果与增长收入")]
bw = (11.93 - 2 * 0.2) / 3
for i, (k, h, d) in enumerate(benefits):
    x = 0.7 + i * (bw + 0.2)
    rrect(s, x, y1, bw, 1.2, PANEL_DARK, 0.8)
    badge(s, x + 0.25, y1 + 0.32, 0.56, k, GOLD, 16)
    text(s, x + 1.0, y1 + 0.18, bw - 1.15, 0.42, h, size=16, color=ICE, bold=True)
    text(s, x + 1.0, y1 + 0.6, bw - 1.15, 0.45, d, size=12, color=MUTED)

# ---------------------------------------------------------------- 11. A & B
s = content_slide("PROJECT 04 · A / B", "冰雪之中的热带乐园与购物天堂")
rrect(s, 0.7, 1.95, 6.6, 4.75)
badge(s, 0.95, 2.18, 0.55, "A", GREEN, 18)
text(s, 1.65, 2.15, 5.5, 0.6, "不去夏威夷，享受热带风情", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
outside = rrect(s, 0.95, 3.0, 2.95, 1.55, RGBColor(0x7F, 0xC8, 0xF8), 0.14, RGBColor(0x7F, 0xC8, 0xF8), 0.1)
text(s, 0.95, 3.0, 2.95, 1.55, [[("室外", {"size": 13, "color": MUTED})], [("冰天雪地", {"size": 22, "bold": True})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=outside, spacing=1.3)
inside = rrect(s, 4.1, 3.0, 2.95, 1.55, GOLD, 0.16, GOLD, 0.1)
text(s, 4.1, 3.0, 2.95, 1.55, [[("场内", {"size": 13, "color": MUTED})], [("四季如春", {"size": 22, "bold": True, "color": GOLD})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=inside, spacing=1.3)
for j, lab in enumerate(["热带风情", "人造沙滩", "温泉海浪"]):
    chip(s, 0.95 + j * 2.07, 4.8, 1.9, 0.5, lab, GREEN, 14)
text(s, 0.95, 5.55, 6.1, 0.9,
     [[("儿童嬉水", {"bold": True, "color": GREEN}), ("  ·  ", {"color": MUTED}), ("成人泡温泉", {"bold": True, "color": GREEN})],
      [("老少皆宜的全天候休闲度假空间", {"size": 13, "color": MUTED})]], size=16, spacing=1.4)

rrect(s, 7.6, 1.95, 5.03, 4.75)
badge(s, 7.85, 2.18, 0.55, "B", GOLD, 18)
text(s, 8.55, 2.15, 4.0, 0.6, "不去拉斯维加斯", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(s, 7.85, 3.05, 4.55, 1.4, [[("豪华", {"color": GOLD})], [("购物天堂", {})]], size=40, bold=True, spacing=1.1)
rect(s, 7.88, 4.6, 0.8, 0.05, GOLD)
text(s, 7.85, 4.8, 4.55, 0.5, "全球最齐备的免税商品", size=18, color=ICE, bold=True)
text(s, 7.85, 5.35, 4.55, 0.9, "一站式汇聚国际品牌，\n为游客提供顶级购物体验", size=13, color=MUTED, spacing=1.4)

# ---------------------------------------------------------------- 12. C food
s = content_slide("PROJECT 04 · C", "不用周游世界，场内即可享受世界美食")
food = [
    ("本地极地海鲜", "世界最绿色、最纯净的海鲜", ["龙虾", "帝王蟹", "北极贝", "北极雪鱼"], TEAL),
    ("品种齐全的中国美食", "汇聚中国各地经典风味", ["北京烤鸭", "上海小笼包", "广东烧鹅", "内蒙烤全羊"], GOLD),
    ("环球异国风味", "足不出“场”尝遍世界", ["日式料理", "泰国美食", "……"], VIOLET),
]
cw, gap = 3.77, 0.31
for i, (h, d, dishes, col) in enumerate(food):
    x = 0.7 + i * (cw + gap)
    rrect(s, x, 1.95, cw, 4.75)
    rect(s, x + 0.02, 1.97, cw - 0.04, 0.08, col)
    text(s, x + 0.3, 2.3, cw - 0.6, 0.5, h, size=20, color=ICE, bold=True)
    text(s, x + 0.3, 2.82, cw - 0.6, 0.4, d, size=13, color=col)
    rect(s, x + 0.3, 3.35, cw - 0.6, 0.015, LINE)
    for j, dish in enumerate(dishes):
        y = 3.6 + j * 0.72
        rrect(s, x + 0.3, y, cw - 0.6, 0.56, col, 0.12, col, 0.3)
        rect(s, x + 0.5, y + 0.22, 0.12, 0.12, col, shape=MSO_SHAPE.OVAL)
        text(s, x + 0.8, y, cw - 1.2, 0.56, dish, size=15, color=ICE, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 13. D & E
s = content_slide("PROJECT 04 · D / E", "极光晚会与冰雪艺术世界")
rrect(s, 0.7, 1.95, 5.95, 4.75, VIOLET, 0.12, VIOLET, 0.05)
badge(s, 0.95, 2.18, 0.55, "D", VIOLET, 18)
text(s, 1.65, 2.15, 4.9, 0.6, "全球独一无二的极光晚会", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
dsteps = [("天上极光", "实时投射到大屏"), ("现场演艺", "音乐、歌舞交相呼应"), ("中国烟花", "与自然界极光交相辉映")]
for j, (h, d) in enumerate(dsteps):
    y = 3.0 + j * 0.92
    rrect(s, 0.95, y, 5.45, 0.74, PANEL, 0.8, VIOLET, 0.15)
    text(s, 1.15, y, 1.5, 0.74, h, size=15, color=GOLD, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 2.65, y, 3.65, 0.74, d, size=14, color=ICE, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.95, 5.85, 5.45, 0.7, "让来自全世界的游客共享艺术盛宴，\n成为世界绝无仅有的自然美景", size=13, color=MUTED, spacing=1.35)

ICEBLUE = RGBColor(0x7F, 0xC8, 0xF8)
rrect(s, 6.95, 1.95, 5.68, 4.75, ICEBLUE, 0.1, ICEBLUE, 0.05)
badge(s, 7.2, 2.18, 0.55, "E", ICEBLUE, 18)
text(s, 7.9, 2.15, 4.6, 0.6, "冰雪构筑的艺术世界", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(s, 7.2, 2.95, 5.2, 0.4, "世界冰雕艺术家精心设计的彩色冰雕：", size=13, color=MUTED)
marks = ["哈尔滨特色城堡", "中国长城", "德国科隆大教堂", "法国凯旋门", "里约耶稣山", "……"]
for j, m in enumerate(marks):
    cx = 7.2 + (j % 2) * 2.65
    cy = 3.45 + (j // 2) * 0.7
    chip(s, cx, cy, 2.5, 0.54, m, ICEBLUE, 14)
text(s, 7.2, 5.65, 5.2, 0.9, "让来看极光的游客，\n更多收获加拿大的美景", size=15, color=ICE, bold=True, spacing=1.35)

# ---------------------------------------------------------------- 14. F economics
s = content_slide("PROJECT 04 · F", "经济效益：旅游带动区域全面发展")
rrect(s, 0.7, 1.95, 5.6, 4.75)
text(s, 0.95, 2.15, 5.1, 0.4, "根据世界旅游权威理论", size=14, color=MUTED)
text(s, 0.95, 2.55, 5.1, 1.3, [[("× 5", {"color": GOLD, "size": 72, "bold": True, "font": "Arial"})]])
text(s, 0.95, 3.85, 5.1, 0.5, "每 100 万美元旅游收入，带动五倍 GDP 增长", size=15, color=ICE, bold=True)
base_y = 6.3
rect(s, 0.95, base_y, 5.1, 0.02, LINE)
b1 = rect(s, 1.4, base_y - 0.38, 1.4, 0.38, TEAL)
b2 = rect(s, 3.6, base_y - 1.75, 1.4, 1.75, GOLD)
gradient(b2, GOLD, RGBColor(0xE8, 0x8A, 0x3C), 90)
text(s, 1.1, base_y - 0.8, 2.0, 0.4, "100 万美元", size=12, color=TEAL, bold=True, align=PP_ALIGN.CENTER)
text(s, 3.3, base_y - 2.15, 2.0, 0.4, "500 万美元", size=12, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
text(s, 1.1, base_y + 0.03, 2.0, 0.3, "旅游收入", size=11, color=MUTED, align=PP_ALIGN.CENTER)
text(s, 3.3, base_y + 0.03, 2.0, 0.3, "GDP 增长", size=11, color=MUTED, align=PP_ALIGN.CENTER)

kpis = [("增加 GDP", GREEN), ("增加就业", TEAL), ("增加城市魅力", VIOLET)]
kw = (6.03 - 2 * 0.2) / 3
for i, (k, col) in enumerate(kpis):
    x = 6.6 + i * (kw + 0.2)
    rrect(s, x, 1.95, kw, 1.35, col, 0.15, col, 0.1)
    text(s, x, 1.95, kw, 1.35, k, size=17, color=ICE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rrect(s, 6.6, 3.55, 6.03, 3.15)
text(s, 6.85, 3.72, 5.5, 0.45, "带动多方面产业发展", size=17, color=ICE, bold=True)
inds = ["飞机", "交通", "旅馆", "商业", "地产", "饮食", "教育"]
for j, ind in enumerate(inds):
    cx = 6.85 + (j % 4) * 1.38
    cy = 4.4 + (j // 4) * 0.8
    chip(s, cx, cy, 1.22, 0.56, ind, TEAL if j % 2 == 0 else GREEN, 15)
text(s, 6.85, 6.0, 5.5, 0.5, "形成“旅游 + ”产业链，放大综合经济效益", size=13, color=MUTED)

# ---------------------------------------------------------------- 15. G & H
s = content_slide("PROJECT 04 · G / H", "投资引入与居民受益")
rrect(s, 0.7, 1.95, 6.9, 4.75)
badge(s, 0.95, 2.18, 0.55, "G", GOLD, 18)
text(s, 1.65, 2.15, 5.8, 0.6, "旅游项目可能引入的投资", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
groups = [
    ("服务与金融", ["服务业", "金融创业"], GOLD),
    ("资源与能源", ["矿产业", "石油、天然气", "绿色电力", "储能"], GREEN),
    ("科技与基础设施", ["AI", "公路交通", "物流"], TEAL),
    ("特色产业", ["淡水渔业", "地热科学（观天象、看极光）"], VIOLET),
]
y = 3.0
for g, its, col in groups:
    text(s, 0.95, y, 1.6, 0.5, g, size=13, color=col, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    cx = 2.6
    for it in its:
        w = 0.3 + 0.19 * len(it)
        chip(s, cx, y + 0.02, w, 0.46, it, col, 13)
        cx += w + 0.12
    y += 0.88

rrect(s, 7.9, 1.95, 4.73, 4.75, GREEN, 0.1, GREEN, 0.05)
badge(s, 8.15, 2.18, 0.55, "H", GREEN, 18)
text(s, 8.85, 2.15, 3.7, 0.6, "当地居民受益情况", size=19, color=ICE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
hs = ["享受所有旅游项目带来的成果", "分享经济增长的收入", "改变寒冷的居住环境", "丰富日常生活"]
for j, h in enumerate(hs):
    y = 3.0 + j * 0.88
    rrect(s, 8.15, y, 4.23, 0.68, PANEL, 0.8, GREEN, 0.15)
    badge(s, 8.3, y + 0.14, 0.4, str(j + 1), GREEN, 13)
    text(s, 8.85, y, 3.45, 0.68, h, size=15, color=ICE, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 16. Closing
s = prs.slides.add_slide(BLANK)
background(s, "cover.jpg")
rect(s, 0, 0, SW, SH, NAVY, 0.35)
text(s, 0, 1.0, SW, 0.4, "SUMMARY", size=14, color=TEAL, bold=True, align=PP_ALIGN.CENTER, font="Arial")
text(s, 0, 1.4, SW, 0.9, "四大项目  协同发展", size=40, color=ICE, bold=True, align=PP_ALIGN.CENTER)
summary = [("01", "教育奠基", "学校教学大楼", GREEN), ("02", "能源保障", "绿色能源与储能", TEAL),
           ("03", "人才资金", "UN 城市与移民基金", VIOLET), ("04", "旅游引擎", "INUVIK 极光世界", GOLD)]
cw, gap = 2.6, 0.3
x0 = (SW - (4 * cw + 3 * gap)) / 2
for i, (n, h, d, col) in enumerate(summary):
    x = x0 + i * (cw + gap)
    rrect(s, x, 2.7, cw, 1.75, NAVY, 0.6, col, 0.1, 1.25)
    text(s, x, 2.82, cw, 0.5, n, size=20, color=col, bold=True, align=PP_ALIGN.CENTER, font="Arial")
    text(s, x, 3.3, cw, 0.5, h, size=20, color=ICE, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, 3.82, cw, 0.4, d, size=12, color=MUTED, align=PP_ALIGN.CENTER)
text(s, 0, 4.9, SW, 0.8, "谢谢观看  THANK YOU", size=28, color=ICE, bold=True, align=PP_ALIGN.CENTER)

prs.save(OUTPUT)
print("saved", OUTPUT, "slides:", len(prs.slides))
