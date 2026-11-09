"""Build the NUNAVIK comprehensive development plan deck from the source PDF content."""
import json
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
PHOTOS = HERE / "assets" / "photos"
OUTPUT = HERE / "努纳维克NUNAVIK地区综合发展规划.pptx"
CREDITS = json.loads((HERE / "photo_credits.json").read_text())

FONT = "Microsoft YaHei"

NAVY = RGBColor(0x0F, 0x2B, 0x46)
BLUE = RGBColor(0x1F, 0x5F, 0x99)
SKY = RGBColor(0x8F, 0xB8, 0xDE)
TEXT = RGBColor(0x33, 0x3D, 0x47)
MUTED = RGBColor(0x6B, 0x77, 0x85)
LIGHT = RGBColor(0xF2, 0xF4, 0xF7)
PALE = RGBColor(0xE4, 0xEC, 0xF4)
RULE = RGBColor(0xD5, 0xDB, 0xE3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
L, R = 0.6, 12.733
CW = R - L
page_no = [0]


def set_fill(shape, color, alpha=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if alpha is not None:
        srgb = shape.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        etree.SubElement(srgb, qn("a:alpha")).set("val", str(int(alpha * 100000)))


def rect(slide, x, y, w, h, color=None, alpha=None, line=None, line_w=0.75, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if color is None:
        shp.fill.background()
    else:
        set_fill(shp, color, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    _drop_theme_style(shp)
    return shp


def _drop_theme_style(shape):
    """Theme style references add shadows in some renderers even with an empty effect list."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def hline(slide, x, y, w, color=RULE, weight=0.75):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(weight)
    _drop_theme_style(c)
    return c


def _apply_font(run, size, color, bold, font=FONT):
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


def text(slide, x, y, w, h, paras, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.2, space_after=0, shape=None, margin=0.04):
    """paras: str | list[str | list[(text, {opts})]]; '\n' in a str starts a new paragraph."""
    tb = shape or slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(margin))
    if isinstance(paras, str):
        paras = [paras]
    paras = [line for p in paras for line in (p.split("\n") if isinstance(p, str) else [p])]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        for txt, opt in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run()
            r.text = txt
            _apply_font(r, opt.get("size", size), opt.get("color", color), opt.get("bold", bold))
    return tb


def bullets(slide, x, y, w, h, items, size=13, color=TEXT, spacing=1.25, gap=6, dot=BLUE):
    return text(slide, x, y, w, h, [[("■  ", {"size": size - 5, "color": dot}), (it, {})] for it in items],
                size=size, color=color, spacing=spacing, space_after=gap)


def photo(slide, key, x, y, w, h, fx=0.5, fy=0.5):
    path = PHOTOS / f"{key}.jpg"
    iw, ih = Image.open(path).size
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    box, img = w / h, iw / ih
    if img > box:
        excess = 1 - box / img
        pic.crop_left, pic.crop_right = excess * fx, excess * (1 - fx)
    else:
        excess = 1 - img / box
        pic.crop_top, pic.crop_bottom = excess * fy, excess * (1 - fy)
    return pic


def caption(slide, x, y, w, label, color=MUTED, align=PP_ALIGN.LEFT):
    text(slide, x, y, w, 0.28, label, size=9, color=color, align=align)


def label_box(slide, x, y, letter, d=0.42, color=NAVY, size=14):
    s = rect(slide, x, y, d, d, color)
    text(slide, x, y, d, d, letter, size=size, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=s, margin=0)
    return s


def chevrons(slide, x, y, w, h, steps, size=13):
    n = len(steps)
    overlap = 0.12
    cw = (w + overlap * (n - 1)) / n
    shades = [NAVY, RGBColor(0x17, 0x45, 0x70), BLUE, RGBColor(0x3A, 0x7C, 0xB8), RGBColor(0x5B, 0x95, 0xC9)]
    for i, st in enumerate(steps):
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        s = rect(slide, x + i * (cw - overlap), y, cw, h, shades[i % len(shades)], shape=kind)
        s.adjustments[0] = 0.3
        text(slide, 0, 0, 0, 0, st, size=size, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, shape=s, margin=0.02)


def footer(slide):
    page_no[0] += 1
    hline(slide, L, 7.0, CW)
    text(slide, L, 7.05, 6, 0.3, "努纳维克 NUNAVIK 地区综合发展规划", size=9, color=MUTED)
    text(slide, R - 1.0, 7.05, 1.0, 0.3, str(page_no[0]), size=9, color=MUTED, align=PP_ALIGN.RIGHT)


def content_slide(title, subtitle=None, section=None):
    s = prs.slides.add_slide(BLANK)
    text(s, L, 0.42, 10.5, 0.6, title, size=24, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        text(s, L, 1.0, 11, 0.36, subtitle, size=13, color=MUTED)
    if section:
        text(s, R - 3.5, 0.55, 3.5, 0.3, section, size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    hline(s, L, 1.455, CW)
    rect(s, L, 1.43, 0.9, 0.05, NAVY)
    footer(s)
    return s


def section_slide(num, title, en, desc, key, cap, fx=0.5, fy=0.5):
    s = prs.slides.add_slide(BLANK)
    photo(s, key, 0, 0, SW, SH, fx, fy)
    rect(s, 0, 0, 5.6, SH, NAVY, 0.9)
    text(s, 0.8, 2.0, 4, 1.0, num, size=54, color=SKY, bold=True)
    rect(s, 0.85, 3.05, 0.8, 0.05, WHITE)
    text(s, 0.8, 3.3, 4.6, 1.2, title, size=30, color=WHITE, bold=True, spacing=1.15)
    text(s, 0.82, 4.5, 4.6, 0.35, en, size=11, color=SKY, bold=True)
    text(s, 0.82, 5.0, 4.4, 1.2, desc, size=13, color=RGBColor(0xD6, 0xE2, 0xEE), spacing=1.45)
    caption(s, 5.8, 7.08, 7.3, cap, WHITE, PP_ALIGN.RIGHT)
    page_no[0] += 1
    return s


# ---------------------------------------------------------------- Cover
s = prs.slides.add_slide(BLANK)
photo(s, "cover_aurora", 5.2, 0, SW - 5.2, SH, 0.45)
rect(s, 0, 0, 5.6, SH, NAVY)
rect(s, 5.6, 0, 0.06, SH, BLUE)
text(s, 0.8, 1.6, 4.6, 0.35, "COMPREHENSIVE DEVELOPMENT PLAN", size=11, color=SKY, bold=True)
text(s, 0.8, 2.1, 4.7, 0.9, "努纳维克 NUNAVIK", size=34, color=WHITE, bold=True)
text(s, 0.8, 2.9, 4.7, 0.8, "地区综合发展规划", size=30, color=WHITE, bold=True)
rect(s, 0.85, 3.9, 0.8, 0.05, SKY)
text(s, 0.8, 4.15, 4.6, 1.6,
     ["学校教学大楼建设项目", "绿色能源项目", "人才与移民基金计划", "“INUVIK 极光世界”旅游项目"],
     size=13, color=RGBColor(0xD6, 0xE2, 0xEE), spacing=1.5)
page_no[0] += 1

# ---------------------------------------------------------------- Contents
s = prs.slides.add_slide(BLANK)
photo(s, "igloo_church", 0, 0, 4.6, SH, 0.42)
caption(s, 0.25, 7.08, 4.2, "因纽维克地标：冰屋教堂", WHITE)
text(s, 5.3, 0.75, 6, 0.6, "目录", size=28, color=NAVY, bold=True)
text(s, 5.3, 1.35, 6, 0.3, "CONTENTS", size=11, color=MUTED, bold=True)
rect(s, 5.32, 1.75, 0.9, 0.05, NAVY)
items = [
    ("01", "学校教学大楼建设项目", "引入中国建设力量 · 成立筹备项目公司 · 前期预算与资金安排"),
    ("02", "绿色能源项目", "风力发电 · 太阳能光伏 · 小型水力发电 · 储能公司"),
    ("03", "人才与移民基金计划", "向 UN 申请设立“NWT”或因纽维克式城市、移民基金"),
    ("04", "“INUVIK 极光世界”旅游项目", "打造北美顶级旅游目的地，带动经济、惠及居民"),
]
for i, (n, t, d) in enumerate(items):
    y = 2.2 + i * 1.15
    text(s, 5.3, y, 1.0, 0.7, n, size=30, color=BLUE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 6.4, y + 0.02, 6.3, 0.4, t, size=18, color=NAVY, bold=True)
    text(s, 6.4, y + 0.44, 6.3, 0.35, d, size=12, color=MUTED)
    if i < 3:
        hline(s, 5.32, y + 0.98, 7.4)
page_no[0] += 1

# ---------------------------------------------------------------- 01
section_slide("01", "学校教学大楼\n建设项目", "SCHOOL BUILDING PROJECT",
              "借助中国建设力量，以项目公司为载体，系统推进高寒地区教学大楼的规划与建设。",
              "college_winter", "因纽维克 Aurora College 校区（冬季）", fx=0.6)

s = content_slide("建设学校教学大楼的项目", "需在前期明确建设力量、实施主体与资金安排三项关键问题", "01 学校教学大楼建设")
blocks = [
    ("A", "是否需要使用中国的建设力量？", "涵盖设计、规划、材料、设备、施工等建设全链条。"),
    ("B", "是否需要成立一个（筹备）项目公司？",
     "由项目公司负责规划、初步设计等前期准备；需赴现场了解冻土施工、高寒地区保温、极昼极夜条件下的采光等问题。"),
    ("C", "前期工作预算与资金安排", "前期工作需要编制预算并做出资金安排，在成立项目公司（B）时设计在股权中。"),
]
for i, (k, h, d) in enumerate(blocks):
    y = 1.8 + i * 1.3
    label_box(s, L, y + 0.04, k)
    text(s, L + 0.6, y, 6.9, 0.45, h, size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, L + 0.6, y + 0.45, 6.9, 0.8, d, size=13, color=TEXT, spacing=1.35)
photo(s, "college_summer", 8.35, 1.8, R - 8.35, 3.75, 0.45)
caption(s, 8.35, 5.58, 4.4, "参考：因纽维克 Aurora College 校区")
text(s, L, 5.72, 4, 0.3, "前期推进路径", size=12, color=MUTED, bold=True)
chevrons(s, L, 6.08, CW, 0.62, ["成立筹备项目公司", "现场调研", "规划与初步设计", "预算与资金安排", "引入中国建设力量"])

# ---------------------------------------------------------------- 02
section_slide("02", "绿色能源项目", "GREEN ENERGY PROJECT",
              "风、光、水多能互补，配套储能系统，为地区发展提供稳定、清洁的电力保障。",
              "wind_field", "冬季风电场（示意）")

s = content_slide("绿色能源项目：多能互补 + 储能", "以风电为优先方向，配套储能解决新能源发电的波动问题", "02 绿色能源")
cards = [
    ("wind_snow", "A", "风力发电", "机会更大一些；需要掌握风力强度、规律等数据。", True),
    ("solar_snow", "B", "太阳能光伏发电", "作为清洁电源的重要补充。", False),
    ("hydro", "C", "小型水力发电", "因地制宜开发小型水电资源。", False),
    ("storage", "D", "建立储能公司", "解决 A、B、C 中电力不稳定及储能问题。", False),
]
cw, gap = (CW - 3 * 0.25) / 4, 0.25
for i, (key, k, h, d, pri) in enumerate(cards):
    x = L + i * (cw + gap)
    rect(s, x, 1.8, cw, 4.0, LIGHT)
    photo(s, key, x, 1.8, cw, 2.05)
    label_box(s, x + 0.2, 4.05, k, 0.38, NAVY if k != "D" else BLUE, 13)
    text(s, x + 0.68, 4.02, cw - 0.8, 0.44, h, size=15, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.2, 4.6, cw - 0.4, 1.1, d, size=12, color=TEXT, spacing=1.35)
    if pri:
        tag = rect(s, x + cw - 1.05, 1.95, 0.9, 0.3, BLUE)
        text(s, 0, 0, 0, 0, "优先方向", size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, shape=tag, margin=0)
band = rect(s, L, 6.05, CW, 0.65, PALE)
text(s, 0, 0, 0, 0, [[("风电 + 光伏 + 小水电", {"bold": True, "color": NAVY}), ("   ▶   ", {"color": BLUE}),
                      ("储能系统", {"bold": True, "color": NAVY}), ("   ▶   ", {"color": BLUE}),
                      ("稳定、清洁、可持续的电力供应", {"bold": True, "color": NAVY})]],
     size=14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=band)

# ---------------------------------------------------------------- 03
section_slide("03", "人才与移民\n基金计划", "TALENT & IMMIGRATION FUND",
              "向 UN 申请设立“NWT”或因纽维克式城市、移民基金，引进人才与国际资金。",
              "un_hq", "纽约联合国总部大楼", fx=0.5, fy=0.3)

s = content_slide("向 UN 申请设立“NWT”或因纽维克式城市、移民基金", "通过移民计划同步解决“人才”与“资金”两大发展瓶颈", "03 人才与移民基金")
photo(s, "citizenship", L, 1.8, 4.7, 4.85, 0.55)
caption(s, L, 6.68, 4.7, "加拿大入籍仪式（示意）")
x0 = 5.75
label_box(s, x0, 1.82, "A")
text(s, x0 + 0.6, 1.78, 6.3, 0.5, "引进人才：当地经济发展的需要", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
fields = ["能源技术", "储能技术", "全球旅游规划设计", "城市管理", "教育及技术培训", "酒店管理", "物流", "金融"]
fw = (R - x0 - 0.2) / 2
for j, f in enumerate(fields):
    fx_ = x0 + (j % 2) * (fw + 0.2)
    fy_ = 2.45 + (j // 2) * 0.52
    rect(s, fx_, fy_, fw, 0.42, LIGHT)
    rect(s, fx_, fy_, 0.06, 0.42, BLUE)
    text(s, fx_ + 0.2, fy_, fw - 0.3, 0.42, f, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
label_box(s, x0, 4.7, "B")
text(s, x0 + 0.6, 4.66, 6.3, 0.5, "引入国际资金：平衡资金缺口", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
flow = [("通过移民", "从国际引入资金"), ("投入建设", "投入到 INUVIK 建设之中"), ("平衡缺口", "弥补地区发展资金缺口")]
bw = (R - x0 - 2 * 0.35) / 3
for j, (h, d) in enumerate(flow):
    bx = x0 + j * (bw + 0.35)
    rect(s, bx, 5.3, bw, 1.35, PALE if j < 2 else NAVY)
    text(s, bx + 0.15, 5.42, bw - 0.3, 0.4, h, size=15, bold=True, color=NAVY if j < 2 else WHITE)
    text(s, bx + 0.15, 5.85, bw - 0.3, 0.7, d, size=12, color=TEXT if j < 2 else PALE, spacing=1.3)
    if j < 2:
        a = rect(s, bx + bw + 0.08, 5.82, 0.2, 0.3, BLUE, shape=MSO_SHAPE.RIGHT_ARROW)

# ---------------------------------------------------------------- 04
section_slide("04", "“INUVIK 极光世界”\n旅游项目", "INUVIK AURORA WORLD",
              "建立北美（至少加拿大）最大、最豪华、最有吸引力的旅游项目。",
              "aurora_teepees", "加拿大北部极光（示意）")

s = content_slide("“INUVIK 极光世界”：五大主题板块",
                  "目标：建立北美（至少加拿大）最大、最豪华、最有吸引力的旅游项目", "04 极光世界旅游项目")
zones = [
    ("tropical", "A", "热带风情乐园", "不去夏威夷", "人造沙滩、温泉海浪，场内四季如春"),
    ("duty_free", "B", "免税购物天堂", "不去拉斯维加斯", "豪华购物，全球最齐备的免税商品"),
    ("king_crab", "C", "环球美食", "不用周游世界", "极地海鲜、中国美食、日料泰餐"),
    ("aurora_lake", "D", "极光晚会", "全球独一无二", "极光投射大屏，音乐歌舞与烟花"),
    ("harbin", "E", "冰雪艺术世界", "世界冰雕名作", "长城、凯旋门、科隆大教堂等"),
]
zw, zg = (CW - 4 * 0.2) / 5, 0.2
for i, (key, k, h, tag, d) in enumerate(zones):
    x = L + i * (zw + zg)
    rect(s, x, 1.75, zw, 3.45, LIGHT)
    photo(s, key, x, 1.75, zw, 1.5)
    label_box(s, x + 0.15, 3.4, k, 0.36, NAVY, 12)
    text(s, x + 0.58, 3.37, zw - 0.65, 0.42, h, size=14, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.15, 3.88, zw - 0.3, 0.3, tag, size=11, color=BLUE, bold=True)
    text(s, x + 0.15, 4.2, zw - 0.3, 0.9, d, size=11, color=TEXT, spacing=1.3)
text(s, L, 5.38, 4, 0.3, "综合效益", size=12, color=MUTED, bold=True)
benefits = [("F", "经济带动", "增加 GDP、就业与城市魅力"), ("G", "投资引入", "服务业、矿产、金融、新能源等"),
            ("H", "居民受益", "共享项目成果与经济增长收入")]
bw = (CW - 2 * 0.2) / 3
for i, (k, h, d) in enumerate(benefits):
    x = L + i * (bw + 0.2)
    rect(s, x, 5.72, bw, 0.98, PALE)
    label_box(s, x + 0.2, 6.0, k, 0.42, BLUE)
    text(s, x + 0.8, 5.8, bw - 0.95, 0.4, h, size=14, color=NAVY, bold=True)
    text(s, x + 0.8, 6.2, bw - 0.95, 0.4, d, size=12, color=TEXT)

# ---------------------------------------------------------------- A / B
s = content_slide("热带风情乐园与免税购物天堂", "外部冰天雪地，场内四季如春", "04 极光世界旅游项目")
hw = (CW - 0.35) / 2
for i, (key, k, h, items) in enumerate([
    ("tropical", "A", "不去夏威夷，享受热带风情",
     ["热带风情、人造沙滩、温泉海浪", "外部冰天雪地，场内四季如春", "儿童嬉水，成人泡温泉"]),
    ("duty_free", "B", "不去拉斯维加斯，享受豪华购物",
     ["豪华购物天堂", "全球最齐备的免税商品"]),
]):
    x = L + i * (hw + 0.35)
    photo(s, key, x, 1.8, hw, 2.85)
    caption(s, x, 4.67, hw, "参考：德国热带岛屿室内度假村" if k == "A" else "参考：机场免税购物区")
    label_box(s, x, 5.05, k)
    text(s, x + 0.6, 5.01, hw - 0.6, 0.5, h, size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, x + 0.6, 5.55, hw - 0.6, 1.3, items, size=13)

# ---------------------------------------------------------------- C
s = content_slide("环球美食：不用周游世界，场内即可享受世界美食", None, "04 极光世界旅游项目")
foods = [
    ("king_crab", "本地极地海鲜", "世界最绿色、最纯净的海鲜", ["龙虾", "帝王蟹", "北极贝", "北极雪鱼"]),
    ("peking_duck", "品种齐全的中国美食", "汇聚中国各地经典风味", ["北京烤鸭", "上海小笼包", "广东烧鹅", "内蒙烤全羊"]),
    ("sushi", "环球异国风味", "日式料理、泰国美食……", ["日式料理", "泰国美食", "更多世界美食"]),
]
fw = (CW - 2 * 0.3) / 3
for i, (key, h, d, items) in enumerate(foods):
    x = L + i * (fw + 0.3)
    photo(s, key, x, 1.75, fw, 2.5)
    rect(s, x, 4.25, fw, 2.45, LIGHT)
    rect(s, x, 4.25, fw, 0.05, NAVY)
    text(s, x + 0.25, 4.42, fw - 0.5, 0.42, h, size=16, color=NAVY, bold=True)
    text(s, x + 0.25, 4.84, fw - 0.5, 0.32, d, size=11, color=BLUE)
    text(s, x + 0.25, 5.3, fw - 0.5, 1.3,
         [[("■  ", {"size": 8, "color": BLUE}), (items[j], {}), ("        ", {}),
           ("■  ", {"size": 8, "color": BLUE}) if j + 1 < len(items) else ("", {}),
           (items[j + 1] if j + 1 < len(items) else "", {})] for j in range(0, len(items), 2)],
         size=13, spacing=1.6)

# ---------------------------------------------------------------- D / E
s = content_slide("极光晚会与冰雪艺术世界", "将自然奇观与文化艺术相结合，打造独一无二的体验", "04 极光世界旅游项目")
hw = (CW - 0.35) / 2
pw = (hw - 0.15) / 2
photo(s, "aurora_lake", L, 1.8, pw, 2.6)
photo(s, "fireworks", L + pw + 0.15, 1.8, pw, 2.6, 0.55)
caption(s, L, 4.42, hw, "极光与烟花（示意）")
label_box(s, L, 4.8, "D")
text(s, L + 0.6, 4.76, hw - 0.6, 0.5, "全球独一无二的极光晚会", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, L + 0.6, 5.3, hw - 0.6, 1.6,
        ["将天上的极光投射到大屏，配合现场音乐、歌舞", "中国烟花与自然界的极光交相辉映", "让全世界游客共享艺术盛宴"], size=13)
x2 = L + hw + 0.35
photo(s, "harbin", x2, 1.8, hw, 2.6, 0.5, 0.6)
caption(s, x2, 4.42, hw, "参考：哈尔滨冰雪大世界")
label_box(s, x2, 4.8, "E")
text(s, x2 + 0.6, 4.76, hw - 0.6, 0.5, "冰雪构筑的艺术世界", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, x2 + 0.6, 5.3, hw - 0.6, 1.6,
        ["世界冰雕艺术家精心设计彩色冰雕", "哈尔滨特色城堡、中国长城、德国科隆大教堂、法国凯旋门、里约耶稣山",
         "让来看极光的游客更多收获加拿大的美景"], size=13)

# ---------------------------------------------------------------- F
s = content_slide("经济效益：旅游带动区域全面发展", "根据世界旅游权威理论，每 100 万美元旅游收入将带动五倍的 GDP 增长", "04 极光世界旅游项目")
rect(s, L, 1.8, 5.6, 4.9, LIGHT)
text(s, L + 0.3, 1.95, 5.0, 0.35, "旅游收入对 GDP 的带动效应（万美元）", size=12, color=MUTED, bold=True)
cd = CategoryChartData()
cd.categories = ["旅游收入", "带动 GDP 增长"]
cd.add_series("金额", (100, 500))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(L + 0.2), Inches(2.35), Inches(5.2), Inches(4.2), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.size = Pt(12)
ch.font.name = FONT
ch.font.color.rgb = TEXT
plot = ch.plots[0]
plot.gap_width = 80
plot.has_data_labels = True
dl = plot.data_labels
dl.font.size = Pt(16)
dl.font.bold = True
dl.font.color.rgb = NAVY
dl.position = XL_LABEL_POSITION.OUTSIDE_END
ser = plot.series[0]
for idx, col in enumerate([SKY, NAVY]):
    pt = ser.points[idx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
va.maximum_scale = 600
va.minimum_scale = 0
ch.category_axis.format.line.color.rgb = RULE
ch.category_axis.tick_labels.font.size = Pt(13)

x2 = 6.55
photo(s, "airport", x2, 1.8, R - x2, 2.15, 0.5, 0.55)
caption(s, x2, 3.97, R - x2, "因纽维克机场：航空与交通将同步受益")
kw = (R - x2 - 2 * 0.15) / 3
for i, k in enumerate(["增加 GDP", "增加就业", "增加城市魅力"]):
    b = rect(s, x2 + i * (kw + 0.15), 4.35, kw, 0.7, NAVY)
    text(s, 0, 0, 0, 0, k, size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)
text(s, x2, 5.25, 5, 0.35, "带动多方面产业发展", size=13, color=NAVY, bold=True)
inds = ["飞机", "交通", "旅馆", "商业", "地产", "饮食", "教育"]
iw = (R - x2 - 6 * 0.1) / 7
for i, ind in enumerate(inds):
    b = rect(s, x2 + i * (iw + 0.1), 5.7, iw, 0.55, PALE)
    text(s, 0, 0, 0, 0, ind, size=13, color=NAVY, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b, margin=0)
text(s, x2, 6.35, R - x2, 0.35, "形成“旅游 +”产业链，放大综合经济效益", size=11, color=MUTED)

# ---------------------------------------------------------------- G / H
s = content_slide("投资引入与居民受益", "旅游项目将吸引多元产业投资，发展成果惠及当地居民", "04 极光世界旅游项目")
label_box(s, L, 1.82, "G")
text(s, L + 0.6, 1.78, 6, 0.5, "旅游项目可能引入的投资", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
rows = [("领域", "投资方向"), ("服务与金融", "服务业、金融创业"), ("资源与能源", "矿产业，石油、天然气，绿色电力，储能"),
        ("科技与基础设施", "AI，公路交通，物流"), ("特色产业", "淡水渔业，地热科学（观天象、看极光）")]
tw = 6.6
tbl = s.shapes.add_table(len(rows), 2, Inches(L), Inches(2.45), Inches(tw), Inches(0.84 * len(rows))).table
tbl.columns[0].width = Inches(1.9)
tbl.columns[1].width = Inches(tw - 1.9)
for r_, row in enumerate(rows):
    tbl.rows[r_].height = Inches(0.84)
    for c_, val in enumerate(row):
        cell = tbl.cell(r_, c_)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if r_ == 0 else (LIGHT if r_ % 2 else WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.15)
        tf = cell.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].text = ""
        run = tf.paragraphs[0].add_run()
        run.text = val
        _apply_font(run, 13, WHITE if r_ == 0 else (NAVY if c_ == 0 else TEXT), r_ == 0 or c_ == 0)

x2 = 7.65
photo(s, "inuvik_street", x2, 1.8, R - x2, 2.3, 0.5, 0.6)
caption(s, x2, 4.12, R - x2, "因纽维克街景")
label_box(s, x2, 4.5, "H")
text(s, x2 + 0.6, 4.46, 4.5, 0.5, "当地居民受益情况", size=16, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
hs = ["享受所有旅游项目带来的成果", "分享经济增长的收入", "改变寒冷的居住环境", "丰富日常生活"]
for j, h in enumerate(hs):
    y = 5.05 + j * 0.42
    text(s, x2, y, 0.4, 0.38, f"{j + 1:02d}", size=13, color=BLUE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x2 + 0.5, y, R - x2 - 0.5, 0.38, h, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- Summary
s = prs.slides.add_slide(BLANK)
photo(s, "delta", 0, 0, SW, SH)
rect(s, 0, 0, SW, SH, NAVY, 0.84)
text(s, 0, 1.0, SW, 0.35, "SUMMARY", size=12, color=SKY, bold=True, align=PP_ALIGN.CENTER)
text(s, 0, 1.4, SW, 0.8, "四大项目  协同发展", size=34, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
rect(s, SW / 2 - 0.4, 2.3, 0.8, 0.05, SKY)
summary = [("01", "教育奠基", "学校教学大楼建设"), ("02", "能源保障", "风光水储多能互补"),
           ("03", "人才资金", "UN 城市与移民基金"), ("04", "旅游引擎", "INUVIK 极光世界")]
sw_, sg = 2.6, 0.3
x0 = (SW - (4 * sw_ + 3 * sg)) / 2
for i, (n, h, d) in enumerate(summary):
    x = x0 + i * (sw_ + sg)
    rect(s, x, 2.85, sw_, 1.8, WHITE, 0.08, SKY, 0.75)
    text(s, x, 3.0, sw_, 0.45, n, size=18, color=SKY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, 3.45, sw_, 0.5, h, size=20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, 4.0, sw_, 0.4, d, size=12, color=PALE, align=PP_ALIGN.CENTER)
text(s, 0, 5.3, SW, 0.7, "谢谢！", size=28, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
caption(s, 5.8, 7.08, 7.3, "麦肯齐河三角洲航拍", WHITE, PP_ALIGN.RIGHT)
page_no[0] += 1

# ---------------------------------------------------------------- Image credits
DESC = {
    "cover_aurora": "封面 · 极光露营", "aurora_teepees": "章节 04 · 极光下的帐篷", "aurora_lake": "湖面极光",
    "igloo_church": "目录 · 因纽维克冰屋教堂", "college_winter": "章节 01 · Aurora College（冬）",
    "college_summer": "Aurora College（夏）", "inuvik_street": "因纽维克街景", "delta": "总结 · 麦肯齐河三角洲",
    "wind_snow": "雪原风机", "wind_field": "章节 02 · 冬季风电场", "solar_snow": "雪地光伏阵列", "hydro": "河道式水电站",
    "storage": "集装箱式储能系统", "un_hq": "章节 03 · 联合国总部", "citizenship": "加拿大入籍仪式",
    "tropical": "德国热带岛屿度假村", "duty_free": "免税购物区", "king_crab": "帝王蟹", "peking_duck": "北京烤鸭",
    "sushi": "日式寿司", "fireworks": "烟花", "harbin": "哈尔滨冰雪大世界", "airport": "因纽维克机场",
}
s = content_slide("图片来源", "本演示文稿图片均来自 Wikimedia Commons，按各自许可协议使用，仅作示意")
half = (len(CREDITS) + 1) // 2
for col in range(2):
    lines = []
    for c in CREDITS[col * half:(col + 1) * half]:
        artist = c["artist"].split("(")[0].strip()[:34] or "Unknown"
        lines.append([(DESC.get(c["key"], c["key"]), {"bold": True, "color": NAVY}),
                      (f"  —  {artist}  ·  {c['license']}", {"color": MUTED})])
    text(s, L + col * 6.15, 1.75, 6.0, 5.1, lines, size=10, spacing=1.15, space_after=5)

prs.save(OUTPUT)
print("saved", OUTPUT, "slides:", len(prs.slides))
