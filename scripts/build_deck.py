#!/usr/bin/env python3
"""Build the Chinese Mission 12458 general-purpose investment deck."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets" / "generated"
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)

W = 13.333
H = 7.5

NAVY = "0E1D2B"
NAVY_2 = "162C3B"
GREEN = "0F5B4C"
TEAL = "2F8B7B"
GOLD = "C69A3A"
SAND = "EFE6D4"
OFFWHITE = "F7F8F5"
PALE_GREEN = "E7F0EB"
PALE_BLUE = "E9F0F4"
TEXT = "17242D"
MUTED = "63727B"
LINE = "D6DEDA"
WHITE = "FFFFFF"
RED = "A44A3F"
FONT = "Microsoft YaHei"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def add_rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str = WHITE,
    line: str | None = None,
    radius: bool = False,
    transparency: int = 0,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.fill.transparency = transparency
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(0.7)
    else:
        shape.line.fill.background()
    if radius:
        shape.adjustments[0] = 0.08
    return shape


def add_text(
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: float = 16,
    color: str = TEXT,
    bold: bool = False,
    font: str = FONT,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin: float = 0.02,
    line_spacing: float = 1.05,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = align
    paragraph.line_spacing = line_spacing
    paragraph.space_after = Pt(0)
    for run in paragraph.runs:
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
    return box


def add_bullets(
    slide,
    items: Iterable[str],
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: float = 14,
    color: str = TEXT,
    bullet_color: str = TEAL,
    gap: float = 6,
    level: int = 0,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = f"•  {item}"
        paragraph.level = level
        paragraph.space_after = Pt(gap)
        paragraph.line_spacing = 1.1
        for run in paragraph.runs:
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.color.rgb = rgb(color)
        if paragraph.runs:
            paragraph.runs[0].font.color.rgb = rgb(bullet_color)
    return box


def add_picture_cover(slide, path: Path, x: float, y: float, w: float, h: float):
    with Image.open(path) as image:
        image_ratio = image.width / image.height
    target_ratio = w / h
    picture = slide.shapes.add_picture(
        str(path), Inches(x), Inches(y), Inches(w), Inches(h)
    )
    if image_ratio > target_ratio:
        visible = target_ratio / image_ratio
        crop = (1 - visible) / 2
        picture.crop_left = crop
        picture.crop_right = crop
    elif image_ratio < target_ratio:
        visible = image_ratio / target_ratio
        crop = (1 - visible) / 2
        picture.crop_top = crop
        picture.crop_bottom = crop
    return picture


def add_header(slide, kicker: str, title: str, subtitle: str | None = None):
    add_text(slide, kicker.upper(), 0.55, 0.30, 5.3, 0.25, size=8.5, color=TEAL, bold=True)
    add_text(slide, title, 0.55, 0.58, 12.1, 0.55, size=26, color=TEXT, bold=True)
    if subtitle:
        add_text(slide, subtitle, 0.57, 1.15, 12.0, 0.34, size=10.5, color=MUTED)
    add_rect(slide, 0.55, 1.52, 12.23, 0.012, fill=LINE)


def add_footer(slide, page: int, source: str | None = None, *, dark: bool = False):
    color = "B9C6CD" if dark else "7A888F"
    if source:
        add_text(slide, source, 0.55, 7.16, 10.9, 0.18, size=6.8, color=color)
    add_text(
        slide,
        f"{page:02d}",
        12.15,
        7.12,
        0.6,
        0.22,
        size=8,
        color=color,
        bold=True,
        align=PP_ALIGN.RIGHT,
    )


def add_metric(
    slide,
    value: str,
    label: str,
    x: float,
    y: float,
    w: float,
    *,
    accent: str = GREEN,
    note: str | None = None,
):
    add_rect(slide, x, y, w, 1.25, fill=WHITE, line=LINE, radius=True)
    add_rect(slide, x, y, 0.08, 1.25, fill=accent)
    add_text(slide, value, x + 0.22, y + 0.17, w - 0.32, 0.43, size=23, color=accent, bold=True)
    add_text(slide, label, x + 0.22, y + 0.65, w - 0.32, 0.28, size=10.5, color=TEXT, bold=True)
    if note:
        add_text(slide, note, x + 0.22, y + 0.95, w - 0.32, 0.18, size=7.2, color=MUTED)


def add_card(
    slide,
    title: str,
    body: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    number: str | None = None,
    fill: str = WHITE,
    accent: str = TEAL,
):
    add_rect(slide, x, y, w, h, fill=fill, line=LINE, radius=True)
    if number:
        add_rect(slide, x + 0.18, y + 0.20, 0.42, 0.42, fill=accent, radius=True)
        add_text(
            slide,
            number,
            x + 0.18,
            y + 0.22,
            0.42,
            0.30,
            size=9,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        title_x = x + 0.72
        title_w = w - 0.9
    else:
        title_x = x + 0.22
        title_w = w - 0.44
    add_text(slide, title, title_x, y + 0.20, title_w, 0.34, size=13, color=TEXT, bold=True)
    add_text(slide, body, x + 0.22, y + 0.74, w - 0.44, h - 0.9, size=9.2, color=MUTED, line_spacing=1.15)


def add_source_chip(slide, text: str, x: float, y: float, w: float):
    add_rect(slide, x, y, w, 0.28, fill=NAVY, radius=True, transparency=10)
    add_text(
        slide,
        text,
        x + 0.08,
        y + 0.045,
        w - 0.16,
        0.16,
        size=6.5,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )


def build() -> Path:
    presentation = Presentation()
    presentation.slide_width = Inches(W)
    presentation.slide_height = Inches(H)
    blank = presentation.slide_layouts[6]

    # 01 Cover
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=NAVY)
    add_picture_cover(slide, ASSETS / "facility_aerial.jpg", 5.0, 0, 8.333, H)
    add_rect(slide, 4.4, 0, 2.6, H, fill=NAVY, transparency=42)
    add_text(slide, "BCQS  |  SPACIOUS CONSTRUCTION", 0.62, 0.52, 4.5, 0.28, size=8.5, color="B8D9D1", bold=True)
    add_text(slide, "MISSION 12458", 0.62, 1.50, 4.25, 0.52, size=31, color=WHITE, bold=True)
    add_text(slide, "综合产业合作项目", 0.62, 2.08, 4.35, 0.58, size=27, color=WHITE, bold=True)
    add_text(
        slide,
        "面向制造、商业、运营与长期开发合作方的项目建议书",
        0.65,
        2.90,
        3.85,
        0.64,
        size=12.5,
        color="D7E3E6",
        line_spacing=1.15,
    )
    add_rect(slide, 0.65, 4.05, 0.72, 0.045, fill=TEAL)
    add_text(
        slide,
        "12458 Dewdney Trunk Road\nMission, British Columbia",
        0.65,
        4.27,
        3.6,
        0.78,
        size=11.5,
        color=WHITE,
        bold=True,
        line_spacing=1.2,
    )
    add_text(slide, "2026 · 合作洽谈版", 0.65, 6.57, 3.0, 0.28, size=9, color="A8BAC4")
    add_text(slide, "概念效果图，仅供合作讨论", 9.70, 7.10, 3.0, 0.18, size=6.5, color=WHITE, align=PP_ALIGN.RIGHT)

    # 02 Opportunity summary
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "01 / Opportunity", "一宗完整资产，分阶段释放产业价值", "一期从CH1商业及轻制造用地启动，长期土地储备保留规划选择。")
    add_metric(slide, "36.9", "总占地 / 英亩", 0.58, 1.82, 2.85, note="约149,329平方米")
    add_metric(slide, "6.0", "CH1一期用地 / 英亩", 3.60, 1.82, 2.85, accent=TEAL, note="约24,281平方米")
    add_metric(slide, "30.9", "RU36长期储备 / 英亩", 6.62, 1.82, 2.85, accent=GOLD, note="当前仍为乡村分区")
    add_metric(slide, "1", "统一产权地块", 9.64, 1.82, 2.85, accent=NAVY_2, note="PID 013-384-228")
    add_rect(slide, 0.58, 3.38, 12.0, 2.72, fill=WHITE, line=LINE, radius=True)
    add_text(slide, "合作命题", 0.88, 3.68, 2.0, 0.30, size=12, color=GREEN, bold=True)
    add_text(
        slide,
        "以清晰的一期开发边界，匹配合作方的真实业务需求；\n以统一总体规划，为后续扩展、配套和资产结构保留弹性。",
        0.88,
        4.10,
        5.7,
        1.25,
        size=18,
        color=TEXT,
        bold=True,
        line_spacing=1.25,
    )
    add_bullets(
        slide,
        [
            "可承接轻型制造、装配、研发、办公、培训、零售展示及住宿等组合。",
            "支持土地交易、长期租赁、定制建设及合资开发等多种合作方式。",
            "所有扩展均以市府用途确认、专业尽调与最终审批为前提。",
        ],
        7.10,
        3.75,
        4.95,
        1.80,
        size=10.7,
        gap=9,
    )
    add_footer(slide, 2, "资料基础：项目方现有商业计划及City of Mission分区文件。")

    # 03 Market context
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "02 / Market Context", "区域增长与供应链需求，为产业项目提供长期支撑")
    add_picture_cover(slide, ASSETS / "vancouver_centerm.jpg", 0.55, 1.80, 5.0, 4.76)
    add_rect(slide, 0.55, 5.44, 5.0, 1.12, fill=NAVY, transparency=12)
    add_text(slide, "3.78M TEUs", 0.82, 5.63, 2.2, 0.38, size=22, color=WHITE, bold=True)
    add_text(slide, "Port of Vancouver 2025集装箱量\n同比增长9%", 2.85, 5.60, 2.35, 0.60, size=10, color=WHITE, bold=True)
    add_card(
        slide,
        "Mission人口与就业增长",
        "现行OCP资料显示，Mission约有46,000名居民，未来20年预计新增约20,000人；规划强调增加本地就业、减少通勤依赖。",
        5.90,
        1.83,
        3.05,
        1.77,
        number="01",
    )
    add_card(
        slide,
        "物流与轻工业进入城市长期战略",
        "Mission OCP明确提及物流及轻工业土地需求，并鼓励沿Lougheed Highway布局仓储、轻制造与物流设施。",
        9.18,
        1.83,
        3.05,
        1.77,
        number="02",
    )
    add_card(
        slide,
        "产业用地稀缺带来择址机会",
        "市府就业用地策略指出：平坦、已服务、交通可达且高于洪泛建设标高的产业用地供应有限。",
        5.90,
        3.86,
        3.05,
        1.77,
        number="03",
    )
    add_card(
        slide,
        "本地租赁参考",
        "Mission市府经济发展网页列示工业空间平均租金约CAD 10–14+/平方英尺，另加附加租金；仅作市场背景，不构成估值或报价。",
        9.18,
        3.86,
        3.05,
        1.77,
        number="04",
    )
    add_text(
        slide,
        "结论：应以“可分期、可定制、审批路径透明”的资产定位进入市场，而不是预设单一业态。",
        5.92,
        6.03,
        6.30,
        0.45,
        size=12,
        color=GREEN,
        bold=True,
    )
    add_footer(
        slide,
        3,
        "来源：Port of Vancouver 2025 Statistics；City of Mission OCP 2025/2026；Employment Lands Strategy；Commercial & Industrial Real Estate。",
    )

    # 04 Location
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "03 / Location", "连接大温、菲沙河谷与美加跨境市场")
    add_picture_cover(slide, ASSETS / "mission_fraser_river.jpg", 8.55, 1.65, 4.23, 5.02)
    add_rect(slide, 8.55, 5.72, 4.23, 0.95, fill=NAVY, transparency=10)
    add_text(slide, "Mission, British Columbia", 8.82, 5.93, 3.5, 0.30, size=13, color=WHITE, bold=True)
    # network schematic
    add_rect(slide, 0.58, 1.82, 7.52, 4.85, fill=WHITE, line=LINE, radius=True)
    nodes = [
        ("温哥华港口", 1.00, 2.55, NAVY),
        ("大温市场", 2.45, 3.70, NAVY_2),
        ("MISSION\n项目", 4.45, 3.02, GREEN),
        ("YXX机场", 6.55, 2.25, TEAL),
        ("美国边境", 6.45, 4.90, GOLD),
        ("Highway 1", 4.10, 5.48, MUTED),
    ]
    lines = [
        (1.62, 2.92, 4.52, 3.43),
        (3.10, 4.05, 4.50, 3.55),
        (5.56, 3.25, 6.66, 2.72),
        (5.50, 3.75, 6.58, 5.05),
        (4.93, 3.86, 4.65, 5.47),
    ]
    for x1, y1, x2, y2 in lines:
        connector = slide.shapes.add_connector(
            1,
            Inches(x1),
            Inches(y1),
            Inches(x2),
            Inches(y2),
        )
        connector.line.color.rgb = rgb("A9BBB5")
        connector.line.width = Pt(2.0)
    for label, x, y, color in nodes:
        add_rect(slide, x, y, 1.22, 0.72, fill=color, radius=True)
        add_text(
            slide,
            label,
            x + 0.05,
            y + 0.14,
            1.12,
            0.42,
            size=9.2,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
    add_text(slide, "MISSION市域参考车程", 0.92, 5.92, 2.15, 0.25, size=9, color=GREEN, bold=True)
    add_text(
        slide,
        "Highway 1 约15分钟  |  Abbotsford机场约25分钟  |  美国边境约25分钟  |  YVR约90分钟",
        0.92,
        6.24,
        6.72,
        0.28,
        size=8.3,
        color=MUTED,
    )
    add_footer(slide, 4, "来源：City of Mission Community Profile 2022。车程为市域参考，项目实际路线与高峰时段需专项核实。")

    # 05 Asset snapshot
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "04 / Asset", "项目资产概况")
    add_picture_cover(slide, ASSETS / "parcel_zoning_map.jpg", 0.58, 1.82, 6.18, 4.90)
    add_source_chip(slide, "项目地块及CH1示意范围", 0.82, 6.20, 2.65)
    labels = [
        ("市政地址", "12458 Dewdney Trunk Road, Mission, BC"),
        ("地块编号", "PID 013-384-228"),
        ("总用地", "约36.9英亩 / 149,329平方米"),
        ("分区构成", "约6.0英亩CH1 + 30.9英亩RU36"),
        ("参考尺度", "约1,280 ft × 1,340 ft"),
        ("产权与边界", "以产权检索、正式测绘及市府记录为准"),
    ]
    start_y = 1.85
    for index, (label, value) in enumerate(labels):
        y = start_y + index * 0.75
        add_text(slide, label, 7.22, y, 1.05, 0.28, size=8.8, color=TEAL, bold=True)
        add_text(slide, value, 8.38, y - 0.01, 4.25, 0.50, size=11.2, color=TEXT, bold=index in (2, 3))
        add_rect(slide, 7.22, y + 0.49, 5.30, 0.012, fill=LINE)
    add_rect(slide, 7.22, 6.25, 5.30, 0.48, fill=SAND, radius=True)
    add_text(
        slide,
        "边界图为项目资料摘录；不能替代测量、产权、环境或用途审查。",
        7.42,
        6.39,
        4.90,
        0.20,
        size=8.2,
        color=TEXT,
        bold=True,
    )
    add_footer(slide, 5, "来源：项目方商业计划、Mission分区图附件（2022）。")

    # 06 Two land classes
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "05 / Land Strategy", "一宗资产，两类土地：先启动、再验证、后扩展")
    add_rect(slide, 0.58, 1.88, 7.15, 1.15, fill=PALE_GREEN, radius=True)
    add_rect(slide, 0.58, 1.88, 1.17, 1.15, fill=GREEN, radius=True)
    add_text(slide, "16.3%", 0.70, 2.20, 0.92, 0.35, size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, "CH1  /  6.0英亩", 1.98, 2.12, 2.55, 0.32, size=15, color=GREEN, bold=True)
    add_text(slide, "一期合作及开发启动区", 4.75, 2.16, 2.55, 0.28, size=10.5, color=TEXT, bold=True)
    add_rect(slide, 0.58, 3.34, 7.15, 1.35, fill=SAND, radius=True)
    add_rect(slide, 0.58, 3.34, 5.98, 1.35, fill=GOLD, radius=True)
    add_text(slide, "83.7%", 0.85, 3.78, 1.0, 0.35, size=16, color=WHITE, bold=True)
    add_text(slide, "RU36  /  30.9英亩", 2.05, 3.70, 2.75, 0.36, size=16, color=WHITE, bold=True)
    add_text(slide, "长期土地储备", 5.05, 3.75, 1.35, 0.28, size=10.8, color=WHITE, bold=True)
    add_card(
        slide,
        "第一阶段｜CH1先行",
        "围绕已列明的主要用途开展市府预咨询、合作方需求匹配、概念设计和投资测算。",
        8.10,
        1.90,
        4.42,
        1.25,
        number="1",
        fill=WHITE,
    )
    add_card(
        slide,
        "第二阶段｜整体总体规划",
        "统筹道路、配套、环境缓冲、访客及运营界面，避免一期方案锁死长期土地选择。",
        8.10,
        3.42,
        4.42,
        1.25,
        number="2",
        fill=WHITE,
    )
    add_card(
        slide,
        "第三阶段｜条件成熟后扩展",
        "RU36范围内任何工业、商业、停车或其他非乡村用途，均需另行完成规划与审批。",
        8.10,
        4.94,
        4.42,
        1.25,
        number="3",
        fill=WHITE,
        accent=GOLD,
    )
    add_rect(slide, 0.58, 5.20, 7.15, 0.92, fill=WHITE, line="E5C9C3", radius=True)
    add_text(slide, "关键边界", 0.85, 5.45, 1.05, 0.26, size=10, color=RED, bold=True)
    add_text(
        slide,
        "30.9英亩RU36不是已获批的产业扩展权；招商材料不得将其视为可立即建设面积。",
        1.88,
        5.40,
        5.45,
        0.38,
        size=10.5,
        color=TEXT,
        bold=True,
    )
    add_footer(slide, 6, "分区面积取自项目资料；最终以市府分区图、产权与测量结果为准。")

    # 07 CH1 planning basis
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "06 / Planning Basis", "CH1一期开发底盘：用途广，但边界必须讲清楚")
    add_metric(slide, "2.0", "基础容积率 FSR", 0.58, 1.82, 2.55, note="理论面积上限≠实际可建面积")
    add_metric(slide, "75%", "最高建筑覆盖率", 3.28, 1.82, 2.55, accent=TEAL, note="仍受退界、停车及环境条件约束")
    add_metric(slide, "14.0 m", "主建筑基础高度", 5.98, 1.82, 2.55, accent=GOLD, note="奖金高度需满足适用条件")
    add_metric(slide, "0.2 ha", "新分割最小地块", 8.68, 1.82, 2.55, accent=NAVY_2, note="不代表本项目建议分割")
    add_rect(slide, 0.58, 3.43, 5.88, 2.48, fill=PALE_GREEN, radius=True)
    add_text(slide, "已列明的主要用途（节选）", 0.86, 3.73, 3.4, 0.32, size=14, color=GREEN, bold=True)
    add_bullets(
        slide,
        [
            "轻型制造",
            "酒店、汽车旅馆",
            "主要/次要零售、食品零售、购物中心",
            "一般办公、医疗办公",
            "学院、大学、教育设施、儿童照护",
            "停车场、车辆租赁、车辆服务、室内娱乐",
        ],
        0.86,
        4.14,
        5.1,
        1.48,
        size=9.7,
        gap=3.5,
    )
    add_rect(slide, 6.72, 3.43, 5.80, 2.48, fill="F9F0EE", radius=True)
    add_text(slide, "需要预先确认的事项", 7.02, 3.73, 3.2, 0.32, size=14, color=RED, bold=True)
    add_bullets(
        slide,
        [
            "独立仓储/履约中心并非CH1明确列出的主要用途。",
            "封闭式储存属于附属用途，须服务于获准的主要用途。",
            "具体生产工序、室外堆放、装卸与车辆待运应逐项解释。",
            "住宅仅可作为商业建筑组成部分并位于首层以上。",
        ],
        7.02,
        4.14,
        4.95,
        1.50,
        size=9.8,
        color=TEXT,
        bullet_color=RED,
        gap=6,
    )
    add_rect(slide, 0.58, 6.18, 11.94, 0.50, fill=NAVY, radius=True)
    add_text(
        slide,
        "招商原则：先取得市府用途解释，再确定租赁、定建或投资条款。",
        0.85,
        6.32,
        11.40,
        0.22,
        size=11,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_footer(slide, 7, "来源：City of Mission Zoning Bylaw 5949-2020，Section 1403（2026年合并版）。")

    # 08 Target sectors
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "07 / Partner Fit", "可招商产业组合：围绕“就业、运营、展示、服务”组织")
    add_picture_cover(slide, ASSETS / "facility_aerial.jpg", 8.92, 1.82, 3.62, 4.95)
    add_rect(slide, 8.92, 5.96, 3.62, 0.81, fill=NAVY, transparency=5)
    add_text(slide, "行业示意图，不代表本项目现状或已获批用途", 9.13, 6.18, 3.18, 0.26, size=7.2, color=WHITE, bold=True)
    cards = [
        ("轻制造与装配", "清洁装配、模块化生产、设备集成、部件测试及相关附属存储。"),
        ("研发与运营中心", "工程研发、软件、采购、售后、区域运营、培训及一般办公。"),
        ("展示与商业服务", "产品展示、零售、体验、客户交付、车辆服务及配套餐饮。"),
        ("教育与短期住宿", "学院、职业培训、到访工程团队酒店/汽车旅馆及会议配套。"),
    ]
    positions = [(0.58, 1.85), (4.60, 1.85), (0.58, 4.10), (4.60, 4.10)]
    for index, ((title, body), (x, y)) in enumerate(zip(cards, positions), start=1):
        add_card(slide, title, body, x, y, 3.65, 1.78, number=f"{index:02d}", fill=WHITE)
    add_footer(
        slide,
        8,
        "用途组合仅作招商筛选框架；任何具体企业工艺与业务模式均须完成分区解释、设计及许可审查。概念图不代表现状或许可结果。",
    )

    # 09 Scenario A
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "08 / Scenario A", "情景A｜轻型制造 + 研发运营中心")
    add_picture_cover(slide, ASSETS / "light_manufacturing.jpg", 0.55, 1.80, 6.55, 4.88)
    add_rect(slide, 0.55, 5.88, 6.55, 0.80, fill=NAVY, transparency=6)
    add_text(slide, "概念效果图｜最终产品、工艺与空间由合作方任务书确定", 0.85, 6.12, 5.95, 0.24, size=8.2, color=WHITE, bold=True)
    add_text(slide, "示意功能总量", 7.48, 1.92, 1.8, 0.28, size=10, color=TEAL, bold=True)
    add_text(slide, "8,000–12,000 m²", 7.45, 2.24, 4.70, 0.48, size=27, color=GREEN, bold=True)
    add_text(slide, "一期概念范围，并非承诺建筑面积", 7.48, 2.78, 4.1, 0.25, size=8.5, color=MUTED)
    program = [
        ("40%", "生产 / 装配"),
        ("20%", "物料 / 附属存储 / 发运"),
        ("10%", "质量 / 检测 / 返修"),
        ("15%", "研发 / 办公 / 培训"),
        ("15%", "公用设施 / 员工空间"),
    ]
    for index, (value, label) in enumerate(program):
        y = 3.28 + index * 0.52
        add_text(slide, value, 7.48, y, 0.60, 0.25, size=10.5, color=TEAL, bold=True)
        add_rect(slide, 8.12, y + 0.04, 2.20 * float(value.strip("%")) / 40, 0.12, fill=TEAL)
        add_text(slide, label, 10.50, y - 0.01, 1.70, 0.28, size=9.2, color=TEXT, bold=True)
    add_rect(slide, 7.45, 6.08, 4.82, 0.60, fill=PALE_GREEN, radius=True)
    add_text(
        slide,
        "单层10,000m²约占6英亩用地41%；实际方案仍需核实退界、停车、装卸、环境和消防。",
        7.68,
        6.22,
        4.36,
        0.34,
        size=8.3,
        color=TEXT,
        bold=True,
    )
    add_footer(slide, 9, "功能比例为通用情景；概念图来自现有项目文件，不代表已建成或已许可。")

    # 10 Scenario B
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "09 / Scenario B", "情景B｜展示体验 + 商业服务 + 培训住宿")
    add_picture_cover(slide, ASSETS / "mixed_use.jpg", 7.32, 1.82, 5.18, 4.88)
    add_rect(slide, 7.32, 5.85, 5.18, 0.85, fill=NAVY, transparency=8)
    add_text(slide, "概念效果图｜建筑选址、规模与用途组合待研究", 7.62, 6.12, 4.58, 0.24, size=8.2, color=WHITE, bold=True)
    add_text(slide, "适合的合作命题", 0.60, 1.92, 2.6, 0.35, size=14, color=GREEN, bold=True)
    items = [
        ("品牌展示与客户体验", "展示、咨询、交付、活动与室内体验形成完整客户界面。"),
        ("培训与人才平台", "职业培训、技术课堂、产品发布与合作方团队驻场。"),
        ("酒店与短期住宿", "服务调试、培训、商务访问及多日客户活动。"),
        ("商业与社区配套", "零售、餐饮、办公与符合条件的首层以上住宅研究。"),
    ]
    for index, (title, body) in enumerate(items):
        y = 2.48 + index * 0.93
        add_rect(slide, 0.62, y, 6.15, 0.72, fill=WHITE, line=LINE, radius=True)
        add_text(slide, f"{index + 1:02d}", 0.83, y + 0.18, 0.46, 0.24, size=9, color=TEAL, bold=True)
        add_text(slide, title, 1.40, y + 0.12, 2.25, 0.28, size=11.2, color=TEXT, bold=True)
        add_text(slide, body, 3.64, y + 0.12, 2.82, 0.40, size=8.6, color=MUTED)
    add_rect(slide, 0.62, 6.28, 6.15, 0.42, fill=SAND, radius=True)
    add_text(slide, "商业、制造与公共体验应分阶段投资，避免在一期内叠加过多运营复杂度。", 0.86, 6.40, 5.65, 0.20, size=8.6, color=TEXT, bold=True)
    add_footer(slide, 10, "用途依据CH1主要用途清单；住宅、酒店、公共活动与制造共存方式须由市府及专业顾问确认。")

    # 11 Site planning principles
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "10 / Site Planning", "场地组织：让运营效率、安全和品牌体验同时成立")
    add_picture_cover(slide, ASSETS / "facility_loading.jpg", 0.55, 1.83, 6.10, 4.84)
    add_rect(slide, 0.55, 5.92, 6.10, 0.75, fill=NAVY, transparency=8)
    add_text(slide, "概念效果图｜装卸、停车及消防布局需专项设计", 0.82, 6.14, 5.48, 0.24, size=8.2, color=WHITE, bold=True)
    principles = [
        ("货运直达", "收货、缓冲、生产、检测和发运形成短链路。"),
        ("人车分流", "员工、访客、货车及应急车辆路径清晰分离。"),
        ("可变空间", "大跨、标准模数与预留机电支持后续工艺变化。"),
        ("前场品牌", "办公、展示、培训及访客界面面向主入口组织。"),
        ("绿色缓冲", "保留景观、雨水与邻接边界缓冲，降低场外影响。"),
        ("分期不互锁", "一期道路和公用设施为后续选择留出接口。"),
    ]
    for index, (title, body) in enumerate(principles):
        col = index % 2
        row = index // 2
        x = 6.95 + col * 2.83
        y = 1.88 + row * 1.55
        add_card(slide, title, body, x, y, 2.58, 1.28, number=f"{index + 1}", fill=OFFWHITE)
    add_footer(slide, 11, "本页为规划原则，不构成总平面、施工图或容量承诺。")

    # 12 Infrastructure / due diligence
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "11 / Readiness", "基础设施与前置验证：把不确定性转化为决策清单")
    columns = [0.60, 3.38, 7.30, 10.25]
    widths = [2.55, 3.65, 2.70, 2.47]
    headers = ["事项", "现有资料", "合作决策所需", "下一步动作"]
    for x, w, title in zip(columns, widths, headers):
        add_rect(slide, x, 1.82, w, 0.52, fill=NAVY)
        add_text(slide, title, x + 0.12, 1.96, w - 0.24, 0.22, size=9.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    rows = [
        ("市政供水", "项目资料提出约400米外接水可能", "容量、路径、地役权、消防流量", "向市府/公用事业单位取得书面条件"),
        ("污水与雨水", "现状资料显示使用化粪系统", "工艺排水、员工负荷、雨洪方案", "完成土建与环境初步研究"),
        ("电力与能源", "现场有电力；产能未核实", "峰值负荷、冗余、充电及变压器", "提交合作方负荷表并申请容量评估"),
        ("道路与出入口", "Dewdney Trunk Road入口条件待核", "货车转弯、视距、日交通量、应急通道", "交通工程及道路主管部门预咨询"),
        ("环境与地形", "历史图示含排水、水道/湿地区域提示", "可开发边界、岩土、生态及洪水条件", "测量、岩土、环境与水文尽调"),
    ]
    y = 2.42
    for row_index, row in enumerate(rows):
        fill = WHITE if row_index % 2 == 0 else "F2F5F2"
        height = 0.78
        for x, w, text in zip(columns, widths, row):
            add_rect(slide, x, y, w, height, fill=fill, line=LINE)
            add_text(slide, text, x + 0.13, y + 0.14, w - 0.26, height - 0.22, size=8.0, color=TEXT, bold=x == columns[0], valign=MSO_ANCHOR.MIDDLE)
        y += height
    add_rect(slide, 0.60, 6.53, 12.12, 0.28, fill=SAND, radius=True)
    add_text(slide, "原则：在签署开发或租赁承诺前，先完成用途、服务容量、出入口和环境四项“停止/继续”判断。", 0.80, 6.59, 11.70, 0.16, size=7.8, color=TEXT, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 12, "现状信息来自项目资料及历史公开挂牌；均须由最新书面资料复核。")

    # 13 Collaboration models
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "12 / Deal Structures", "合作方式：从资产交易到定制运营，按风险与控制权匹配")
    models = [
        ("土地/资产交易", "适合希望直接持有并主导开发的合作方。", "合作方：收购与开发\n项目方：产权、资料与交割支持"),
        ("长期土地租赁", "降低前期土地资金占用，锁定长期运营地点。", "合作方：建设及运营\n项目方：土地与长期租约"),
        ("定制建设出租", "按租户任务书设计建设，形成长期租赁资产。", "合作方：规格、租期与信用\n项目方：融资、建设与交付"),
        ("合资项目公司", "土地、资金、开发与运营能力按协议组合。", "双方：资本与治理\n专业团队：设计、成本与施工"),
    ]
    for index, (title, body, roles) in enumerate(models):
        x = 0.58 + index * 3.08
        add_rect(slide, x, 1.90, 2.82, 4.52, fill=OFFWHITE, line=LINE, radius=True)
        add_rect(slide, x, 1.90, 2.82, 0.62, fill=[GREEN, TEAL, NAVY_2, GOLD][index], radius=True)
        add_text(slide, f"{index + 1:02d}", x + 0.18, 2.05, 0.45, 0.25, size=9, color=WHITE, bold=True)
        add_text(slide, title, x + 0.60, 2.00, 1.95, 0.31, size=12.5, color=WHITE, bold=True)
        add_text(slide, "适用场景", x + 0.22, 2.80, 1.25, 0.24, size=8.5, color=TEAL, bold=True)
        add_text(slide, body, x + 0.22, 3.15, 2.38, 0.84, size=10.2, color=TEXT, bold=True, line_spacing=1.15)
        add_rect(slide, x + 0.22, 4.15, 2.38, 0.012, fill=LINE)
        add_text(slide, "核心责任", x + 0.22, 4.42, 1.25, 0.24, size=8.5, color=TEAL, bold=True)
        add_text(slide, roles, x + 0.22, 4.78, 2.38, 0.92, size=9.2, color=MUTED, line_spacing=1.25)
        add_text(slide, "价格 / 租金 / 回报", x + 0.22, 5.88, 2.38, 0.20, size=8.2, color=TEXT, bold=True)
        add_text(slide, "在用途与范围确认后形成Term Sheet", x + 0.22, 6.12, 2.38, 0.24, size=7.8, color=MUTED)
    add_footer(slide, 13, "合作结构需结合税务、融资、责任隔离、许可与退出机制另行设计。")

    # 14 Investment framework
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "13 / Investment", "投资框架：先锁定业务边界，再形成可融资成本")
    stages = [
        ("01", "土地与交易", "价格、税费、产权结构、土地投入或租赁条件"),
        ("02", "场地与基础设施", "土方、道路、排水、供水、供电、消防及环境措施"),
        ("03", "建筑与专业系统", "主体、围护、暖通、电气、消防、办公及员工设施"),
        ("04", "工艺与运营", "设备、自动化、软件、库存、认证、人员与营运资金"),
        ("05", "融资与风险", "预备费、价格上涨、利息、进度、租户信用及退出安排"),
    ]
    for index, (num, title, body) in enumerate(stages):
        y = 1.90 + index * 0.93
        add_rect(slide, 0.62, y, 7.36, 0.70, fill=WHITE, line=LINE, radius=True)
        add_rect(slide, 0.62, y, 0.72, 0.70, fill=[GREEN, TEAL, NAVY_2, GOLD, MUTED][index], radius=True)
        add_text(slide, num, 0.74, y + 0.20, 0.46, 0.22, size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, title, 1.58, y + 0.15, 1.78, 0.30, size=11.5, color=TEXT, bold=True)
        add_text(slide, body, 3.35, y + 0.15, 4.30, 0.34, size=8.7, color=MUTED)
    add_rect(slide, 8.35, 1.90, 4.17, 4.42, fill=NAVY, radius=True)
    add_text(slide, "现有金额口径", 8.72, 2.24, 2.2, 0.30, size=13, color="A9D6CB", bold=True)
    add_text(slide, "CAD 13.65M", 8.70, 2.82, 3.25, 0.52, size=27, color=WHITE, bold=True)
    add_text(slide, "项目资料中的土地预算/价值口径", 8.72, 3.40, 3.15, 0.28, size=9, color="D1DDE2")
    add_rect(slide, 8.72, 3.92, 3.45, 0.012, fill="5C7280")
    add_bullets(
        slide,
        [
            "不是现行挂牌价。",
            "不是独立第三方估值。",
            "不得直接等同合作交易价格。",
            "最终应由最新估值、尽调及商业条款支持。",
        ],
        8.72,
        4.22,
        3.32,
        1.45,
        size=9.4,
        color=WHITE,
        bullet_color="A9D6CB",
        gap=7,
    )
    add_rect(slide, 0.62, 6.48, 11.90, 0.28, fill=SAND, radius=True)
    add_text(slide, "招商阶段不承诺完整项目总投资；合作方任务书确定后，才能编制可比较、可融资、可追责的成本计划。", 0.84, 6.54, 11.45, 0.16, size=7.8, color=TEXT, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 14, "金额来源：项目方2026年商业计划及EV建议书；本页不构成估值、报价或投资回报承诺。")

    # 15 Roadmap
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "14 / Roadmap", "实施路径：以五个决策门控制投入")
    road = [
        ("需求定义", "合作方业务、工艺、面积、流量、投运目标", "输出：项目任务书"),
        ("预咨询与尽调", "用途解释、产权、测量、环境、服务容量、交通", "输出：可行性红旗清单"),
        ("概念与商务", "总图、功能、成本、租赁/投资结构、Term Sheet", "输出：概念方案与商业条款"),
        ("审批与融资", "顾问设计、市府申请、合同、融资与采购策略", "输出：可执行投资决定"),
        ("建设与投运", "施工、设备、调试、验收、培训及运营移交", "输出：合规投运资产"),
    ]
    x_positions = [0.58, 3.08, 5.58, 8.08, 10.58]
    for index, ((title, body, output), x) in enumerate(zip(road, x_positions), start=1):
        add_rect(slide, x, 2.24, 2.12, 3.18, fill=OFFWHITE, line=LINE, radius=True)
        add_rect(slide, x + 0.72, 1.74, 0.68, 0.68, fill=[GREEN, TEAL, NAVY_2, GOLD, MUTED][index - 1], radius=True)
        add_text(slide, str(index), x + 0.84, 1.93, 0.44, 0.24, size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, title, x + 0.18, 2.65, 1.76, 0.36, size=12, color=TEXT, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, body, x + 0.22, 3.28, 1.68, 0.82, size=8.3, color=MUTED, align=PP_ALIGN.CENTER, line_spacing=1.15)
        add_rect(slide, x + 0.22, 4.27, 1.68, 0.012, fill=LINE)
        add_text(slide, output, x + 0.22, 4.55, 1.68, 0.45, size=8.0, color=GREEN, bold=True, align=PP_ALIGN.CENTER)
        if index < len(road):
            add_text(slide, "→", x + 2.14, 3.52, 0.34, 0.30, size=16, color="A8B6B1", bold=True, align=PP_ALIGN.CENTER)
    add_rect(slide, 0.58, 5.95, 12.12, 0.72, fill=NAVY, radius=True)
    add_text(slide, "停止/继续判断", 0.88, 6.16, 1.65, 0.26, size=10.5, color="A9D6CB", bold=True)
    add_text(slide, "每一阶段只在用途、服务容量、成本和合作条款达到共同门槛后进入下一阶段。", 2.52, 6.14, 9.76, 0.30, size=11, color=WHITE, bold=True)
    add_footer(slide, 15, "实际审批及建设周期取决于用途、范围、顾问研究、市府流程、融资与市场条件。")

    # 16 Team
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "15 / Delivery Team", "本地项目协调与建设交付能力")
    add_picture_cover(slide, ASSETS / "facility_entry.jpg", 0.58, 1.82, 4.34, 4.88)
    add_rect(slide, 0.58, 5.75, 4.34, 0.95, fill=NAVY, transparency=8)
    add_text(slide, "Jack Chen, P.Eng. (F), PQS", 0.85, 5.96, 3.70, 0.30, size=12, color=WHITE, bold=True)
    add_text(slide, "工程、工料测量与施工管理", 0.85, 6.34, 3.70, 0.22, size=8, color="D1DDE2")
    roles = [
        ("土地业主 / 项目公司", "场地安排、产权交易或土地投入、本地项目架构及尽调协调。"),
        ("BCQS / Spacious Construction", "可行性、成本规划、顾问协调、招标、施工管理及竣工移交支持。"),
        ("SENMY Canada", "在选定设计与规格范围内提供建筑材料采购及配套供应支持。"),
        ("合作方及专业团队", "业务任务书、运营系统、资本、独立顾问、许可设计与长期运营。"),
    ]
    for index, (title, body) in enumerate(roles):
        y = 1.90 + index * 1.12
        add_rect(slide, 5.28, y, 7.25, 0.88, fill=WHITE, line=LINE, radius=True)
        add_text(slide, f"{index + 1:02d}", 5.52, y + 0.28, 0.42, 0.22, size=8.8, color=TEAL, bold=True)
        add_text(slide, title, 6.08, y + 0.16, 2.28, 0.28, size=10.8, color=TEXT, bold=True)
        add_text(slide, body, 8.40, y + 0.14, 3.76, 0.48, size=8.5, color=MUTED)
    add_rect(slide, 5.28, 6.50, 7.25, 0.22, fill=PALE_GREEN, radius=True)
    add_footer(slide, 16, "团队职责以最终委任、合同和专业责任范围为准；概念图仅作项目愿景表达。")

    # 17 Partner inputs
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "16 / Partner Brief", "合作方需要带来的信息：越具体，方案越可执行")
    inputs = [
        ("业务模型", "拟开展的主要业务、客户、收入来源及运营边界"),
        ("空间任务", "建筑面积、净高、柱网、装卸、停车及户外作业"),
        ("工艺与设备", "生产步骤、设备重量、电力、通风、消防及危险物"),
        ("流量与交通", "员工、访客、货车类型、日均/峰值车次与班次"),
        ("时间目标", "选址决策、许可、建设、设备进场与投运节点"),
        ("商业条件", "租期、资本投入、信用支持、治理及退出偏好"),
        ("合规要求", "行业许可、质量体系、数据、安保及保险标准"),
        ("扩展路线", "一期规模、产能爬坡、第二阶段及土地选择权"),
    ]
    for index, (title, body) in enumerate(inputs):
        col = index % 4
        row = index // 4
        x = 0.58 + col * 3.08
        y = 1.87 + row * 2.15
        add_card(slide, title, body, x, y, 2.82, 1.80, number=f"{index + 1:02d}", fill=OFFWHITE)
    add_rect(slide, 0.58, 6.34, 12.02, 0.48, fill=NAVY, radius=True)
    add_text(slide, "建议首次技术会即使用一页需求表统一口径，避免先画方案、后改业务。", 0.87, 6.48, 11.45, 0.22, size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 17)

    # 18 Decision questions
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=OFFWHITE)
    add_header(slide, "17 / Decision Agenda", "首次面谈建议聚焦八个问题")
    questions = [
        "合作方希望在加拿大建立什么业务能力？",
        "为什么需要Mission，而不是更靠近温哥华的节点？",
        "一期必须具备哪些空间与基础设施？",
        "哪些工序或业务可能触发更严格用途解释？",
        "合作方更看重所有权、现金流还是运营控制权？",
        "一期投资由谁承担，信用支持如何安排？",
        "RU36长期土地对合作方是否真正有价值？",
        "双方以什么条件决定进入下一阶段？",
    ]
    for index, question in enumerate(questions):
        col = index % 2
        row = index // 2
        x = 0.60 + col * 6.05
        y = 1.86 + row * 1.17
        add_rect(slide, x, y, 5.72, 0.88, fill=WHITE, line=LINE, radius=True)
        add_rect(slide, x + 0.18, y + 0.20, 0.48, 0.48, fill=GREEN if col == 0 else TEAL, radius=True)
        add_text(slide, f"{index + 1}", x + 0.18, y + 0.32, 0.48, 0.20, size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, question, x + 0.86, y + 0.20, 4.52, 0.44, size=10.8, color=TEXT, bold=True, valign=MSO_ANCHOR.MIDDLE)
    add_footer(slide, 18, "本页可直接作为保密技术研讨会或现场考察的会议议程。")

    # 19 Next step / contact
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=NAVY)
    add_picture_cover(slide, ASSETS / "factory_gallery.jpg", 7.10, 0, 6.233, H)
    add_rect(slide, 6.55, 0, 1.45, H, fill=NAVY, transparency=35)
    add_text(slide, "NEXT STEP", 0.65, 0.63, 2.0, 0.25, size=9, color="A9D6CB", bold=True)
    add_text(slide, "从联合场地与业务评估开始", 0.65, 1.30, 5.85, 0.58, size=27, color=WHITE, bold=True)
    add_text(
        slide,
        "建议下一步：保密技术研讨  ·  现场考察  ·  市府预咨询",
        0.67,
        2.08,
        5.65,
        0.38,
        size=12,
        color="D7E3E6",
        bold=True,
    )
    steps = [
        ("01", "确认合作方任务书"),
        ("02", "核实土地与基础设施适配"),
        ("03", "形成概念方案、成本与合作条款"),
    ]
    for index, (num, label) in enumerate(steps):
        y = 2.92 + index * 0.72
        add_rect(slide, 0.68, y, 0.48, 0.48, fill=TEAL, radius=True)
        add_text(slide, num, 0.68, y + 0.12, 0.48, 0.18, size=7.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, label, 1.36, y + 0.10, 4.75, 0.25, size=11, color=WHITE, bold=True)
    add_rect(slide, 0.68, 5.34, 5.78, 0.012, fill="506674")
    add_text(slide, "Jack Chen, P.Eng. (F), PQS", 0.68, 5.62, 4.9, 0.32, size=14, color=WHITE, bold=True)
    add_text(slide, "778-896-8116  |  INFO.BCQS@GMAIL.COM", 0.68, 6.10, 5.4, 0.28, size=10.5, color="D7E3E6", bold=True)
    add_text(slide, "Spacious Construction Corp.  |  BCQS  |  www.scco.ca  |  www.bcqs.ca", 0.68, 6.55, 5.82, 0.25, size=8.4, color="A8BAC4")
    add_footer(slide, 19, "概念效果图，仅供合作讨论。", dark=True)

    # 20 Sources and disclaimer
    slide = presentation.slides.add_slide(blank)
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_header(slide, "Appendix", "主要资料来源与免责声明")
    sources_left = [
        "1. City of Mission Zoning Bylaw 5949-2020，Section 1403 CH1/CH2（2026合并版）",
        "2. City of Mission Official Community Plan 6350-2025 / 6405-2026相关公开资料",
        "3. City of Mission Employment Lands Strategy 2021–2041",
        "4. City of Mission Community Profile 2022",
        "5. City of Mission Commercial & Industrial Real Estate网页",
        "6. Port of Vancouver 2025 Statistics Overview及2025年度公开数据",
    ]
    sources_right = [
        "7. 12458 Subdivision Mission Business Plan CN 2026 Final（项目方文件）",
        "8. Mission EV Assembly Proposal CN 2026（项目方文件）",
        "9. 萝莉苑项目规划设计交流两方案对比（2024概念资料）",
        "10. Wikimedia Commons：FraserRiver.jpg；Vancouver Centerm Terminal 2022",
    ]
    add_rect(slide, 0.60, 1.82, 6.08, 3.18, fill=OFFWHITE, line=LINE, radius=True)
    add_text(slide, "官方与市场资料", 0.90, 2.12, 2.30, 0.30, size=13, color=GREEN, bold=True)
    add_bullets(slide, sources_left, 0.90, 2.58, 5.45, 2.10, size=8.5, gap=6)
    add_rect(slide, 6.93, 1.82, 5.82, 3.18, fill=OFFWHITE, line=LINE, radius=True)
    add_text(slide, "项目资料与图片", 7.23, 2.12, 2.30, 0.30, size=13, color=GREEN, bold=True)
    add_bullets(slide, sources_right, 7.23, 2.58, 5.15, 1.85, size=8.5, gap=9)
    add_rect(slide, 0.60, 5.30, 12.15, 1.32, fill=NAVY, radius=True)
    add_text(slide, "免责声明", 0.90, 5.61, 1.15, 0.26, size=10.5, color="A9D6CB", bold=True)
    add_text(
        slide,
        "本文件用于初步招商与合作讨论，不构成要约、估值、规划许可、工程设计、融资承诺或投资回报保证。面积、分区、边界、产权、环境、服务容量、成本、时间及允许用途均应由市府、产权文件、测量及独立专业顾问最终确认。概念图片不代表现状、审批结果或交付标准。",
        2.05,
        5.52,
        10.25,
        0.72,
        size=8.6,
        color=WHITE,
        line_spacing=1.18,
    )
    add_footer(slide, 20, "资料核阅日期：2026年11月。")

    presentation.core_properties.title = "Mission 12458 综合产业合作项目"
    presentation.core_properties.subject = "通用招商合作建议书"
    presentation.core_properties.author = "BCQS / Spacious Construction Corp."
    presentation.core_properties.keywords = "Mission, British Columbia, investment, industrial, commercial"
    output = OUTPUT / "Mission_12458_综合产业合作招商书_2026.pptx"
    presentation.save(output)
    return output


if __name__ == "__main__":
    print(build())
