#!/usr/bin/env python3
"""HR API, 1 декабря 2026. 14 слайдов 16:9.

Цифры, сроки и статусы — только из исходников:
оценка 360 (сентябрь 2026), CV-сервис, Политех / цифровые сотрудники,
бонусные письма, AI Challenge. Не выдумывать экономию и архитектуру.
"""

from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
ASSETS = ROOT / "assets"
OUT_PPTX = REPO / "HR_API_KRIT_Nikisheva_2026.pptx"
OUT_NOTES = REPO / "Speaker_Notes.md"

W = Inches(13.333)
H = Inches(7.5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1E, 0x1E, 0x1E)
MUTED = RGBColor(0x77, 0x77, 0x77)
ACCENT = RGBColor(0x38, 0x9B, 0xFF)
LINE = RGBColor(0xE0, 0xE3, 0xE8)
SOFT = RGBColor(0xF7, 0xF9, 0xFB)
BLUE_SOFT = RGBColor(0xEE, 0xF6, 0xFF)
NAVY = RGBColor(0x0B, 0x2A, 0x4A)
NAVY_MID = RGBColor(0x14, 0x3D, 0x66)
# Как в «Цифровые сотрудники» и AI Challenge: typeface Uncage, текст Arial.
# На этом Mac PowerPoint инвертирует начертания семейства UNCAGE:
# Uncage + b=0 → Bold, Uncage + b=1 → Regular. Regular — нужный фирменный вид.
HEAD = "Uncage"
BODY = "Arial"
HEAD_AS_REGULAR = True
# Сдержанная шкала: Uncage только на заголовках, Arial — всё остальное.
T_DISPLAY = 28
T_TITLE = 22
T_SUB = 20
T_LEAD = 16
T_BODY = 15
T_LABEL = 14
T_KICKER = 12
T_CAPTION = 11
T_KPI = 36
FOOT = "КРИТ · HR API · 1 декабря 2026"
LOGO_GREY = str(ASSETS / "logo-grey.png")
LOGO_WHITE = str(ASSETS / "logo-white.png")
VERTIK = str(ASSETS / "vertik.png")
SPEAKER = str(ASSETS / "speaker.jpg")
QR = str(ASSETS / "qr-nikisheva.png")
SHOT_360_PROCESS = str(ASSETS / "360-process.png")
SHOT_360_DASH = str(ASSETS / "360-dashboard.png")
SHOT_BONUS = str(ASSETS / "bonus-goals-crop.png")
SHOT_CV_REPORT = str(ASSETS / "cv-resume-report.png")

NOTES = {}


def set_run(run, text, size, color, font=BODY, bold=False):
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = font
    if font == HEAD:
        bold = HEAD_AS_REGULAR
    run.font.bold = bold
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "ru-RU")
    rPr.set("b", "1" if bold else "0")
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", font)


def _anchor(body, anchor, pad):
    body.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    if not pad:
        for attr in ("lIns", "tIns", "rIns", "bIns"):
            body.set(attr, "0")


def add_box(slide, l, t, w, h, fill, line=None, line_w=Pt(0.75)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = line_w
    return sh


def add_round(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.adjustments[0] = 0.08
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    return sh


def add_bar(slide, l, t, w, fill=ACCENT, h=Pt(3)):
    return add_box(slide, l, t, w, h, fill)


def add_text(
    slide, l, t, w, h, text, size, color, font=BODY, bold=False,
    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, pad=True, spacing=1.0,
):
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _anchor(tf._txBody.bodyPr, anchor, pad)
    for i, part in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = spacing
        run = p.add_run()
        set_run(run, part, size, color, font=font, bold=bold)
    return sh


def add_lines(slide, l, t, w, h, lines, size, color, font=BODY, space=6, bold=False, spacing=1.0):
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _anchor(tf._txBody.bodyPr, MSO_ANCHOR.TOP, True)
    for i, text in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(space)
        p.line_spacing = spacing
        run = p.add_run()
        set_run(run, text, size, color, font=font, bold=bold)
    return sh


def add_rich(slide, l, t, w, h, paragraphs, anchor=MSO_ANCHOR.TOP, pad=True):
    """paragraphs: list of dicts {align, space, runs:[(text,size,color,font,bold)]}"""
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _anchor(tf._txBody.bodyPr, anchor, pad)
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get("align", PP_ALIGN.LEFT)
        p.space_before = Pt(para.get("before", 0))
        p.space_after = Pt(para.get("after", 0))
        p.line_spacing = para.get("spacing", 1.0)
        for text, size, color, font, bold in para["runs"]:
            run = p.add_run()
            set_run(run, text, size, color, font=font, bold=bold)
    return sh


def notes(slide, key, text):
    NOTES[key] = text
    slide.notes_slide.notes_text_frame.text = text


def blank(prs, fill=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(slide, 0, 0, W, H, fill)
    return slide


def logo_grey(slide):
    slide.shapes.add_picture(LOGO_GREY, Inches(11.95), Inches(0.18), height=Inches(0.42))


def footer(slide, num, color=ACCENT, muted=MUTED):
    add_text(slide, Inches(0.55), Inches(7.16), Inches(9.2), Inches(0.22), FOOT, T_CAPTION, color, pad=False)
    add_text(slide, Inches(11.4), Inches(7.16), Inches(1.4), Inches(0.22), num, T_CAPTION, muted, align=PP_ALIGN.RIGHT, pad=False)


def section(slide, kicker, title, num, title_size=None, title_w=Inches(10.7)):
    title_size = T_TITLE if title_size is None else title_size
    lines = title.count("\n") + 1
    title_h = Inches(title_size * 1.28 / 72 * lines + 0.02)
    add_text(slide, Inches(0.55), Inches(0.22), Inches(10.2), Inches(0.24), kicker, T_KICKER, ACCENT, pad=False)
    add_text(
        slide, Inches(0.55), Inches(0.46), title_w, title_h,
        title, title_size, INK, font=HEAD, pad=False, spacing=0.98,
    )
    logo_grey(slide)
    top = Inches(0.46) + title_h + Inches(0.22)
    add_bar(slide, Inches(0.55), top - Inches(0.12), Inches(1.45), h=Pt(2.5))
    footer(slide, num)
    return top


def place_image(slide, path, l, t, w, h, mode="cover", valign="center"):
    with Image.open(path) as im:
        iw, ih = im.size
    box_ratio = w / h
    img_ratio = iw / ih
    pic = slide.shapes.add_picture(str(path), l, t, width=w, height=h)
    try:
        pic.shadow.inherit = False
    except Exception:
        pass
    if mode == "contain":
        if img_ratio > box_ratio:
            pic.height = int(w / img_ratio)
            pic.top = t + (h - pic.height) // 2
        else:
            pic.width = int(h * img_ratio)
            pic.left = l + (w - pic.width) // 2
            pic.height = h
        return pic
    if img_ratio > box_ratio:
        keep = box_ratio / img_ratio
        side = (1 - keep) / 2
        pic.crop_left = side
        pic.crop_right = side
    else:
        keep = img_ratio / box_ratio
        spare = 1 - keep
        if valign == "top":
            pic.crop_top = spare * 0.08
            pic.crop_bottom = spare - pic.crop_top
        else:
            pic.crop_top = spare / 2
            pic.crop_bottom = spare / 2
    return pic


def outline(slide, l, t, w, h, color=LINE, weight=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    sh.fill.background()
    sh.line.color.rgb = color
    sh.line.width = Pt(weight)
    return sh


def arrow(slide, l, t, w=Inches(0.28), h=Inches(0.18), fill=ACCENT):
    sh = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, l, t, w, h)
    sh.shadow.inherit = False
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def oval(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    sh.shadow.inherit = False
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    return sh


def connector(slide, x1, y1, x2, y2, color=LINE):
    sh = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    sh.line.color.rgb = color
    sh.line.width = Pt(1.25)
    return sh


def pill(slide, l, t, w, h, text, fill, color, size=13):
    add_round(slide, l, t, w, h, fill)
    add_text(slide, l, t, w, h, text, size, color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=False)


# ---------------------------------------------------------------------------
# 01 cover
# ---------------------------------------------------------------------------
def slide_cover(prs):
    s = blank(prs)
    builder = s.shapes.build_freeform(Inches(9.85), Inches(0))
    builder.add_line_segments([(W, Inches(0)), (W, H), (Inches(7.55), H)], close=True)
    geo = builder.convert_to_shape()
    geo.shadow.inherit = False
    geo.fill.solid()
    geo.fill.fore_color.rgb = ACCENT
    geo.line.fill.background()
    s.shapes.add_picture(LOGO_WHITE, Inches(11.88), Inches(0.28), height=Inches(0.46))
    add_box(s, Inches(0.62), Inches(0.46), Inches(1.55), Inches(0.32), BLUE_SOFT)
    add_text(s, Inches(0.62), Inches(0.46), Inches(1.55), Inches(0.32), "HR API  ·  КРИТ", T_CAPTION, ACCENT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, pad=False)
    add_text(
        s, Inches(0.62), Inches(1.18), Inches(8.7), Inches(2.2),
        "HR сам себе\nразработчик",
        T_DISPLAY, INK, font=HEAD, spacing=0.94, pad=False,
    )
    add_box(s, Inches(0.62), Inches(3.52), Inches(1.45), Pt(2.5), ACCENT)
    add_text(
        s, Inches(0.62), Inches(3.78), Inches(8.3), Inches(0.9),
        "Как мы создаём и внедряем\nИИ-инструменты в КРИТ",
        T_LEAD, INK, spacing=1.12, pad=False,
    )
    add_text(
        s, Inches(10.15), Inches(5.92), Inches(2.85), Inches(0.9),
        "1 декабря\n2026",
        T_LEAD, WHITE, pad=False, spacing=1.12,
    )
    s.shapes.add_picture(VERTIK, Inches(0.55), Inches(5.78), height=Inches(1.18))
    add_text(s, Inches(1.92), Inches(6.05), Inches(5.4), Inches(0.36), "Елена Никишева", T_SUB, INK, font=HEAD, pad=False)
    add_text(s, Inches(1.92), Inches(6.44), Inches(5.4), Inches(0.28), "COO КРИТ", T_BODY, MUTED, pad=False)
    notes(
        s, "01",
        "0:40. Не зачитывайте заголовок.\n\n"
        "Начните с рабочей правды. Раньше, чтобы сдвинуть HR-процесс, нужны были подрядчик, бюджет, техническое задание и месяцы ожидания. "
        "Сегодня я расскажу, как HR-команда в КРИТ перестала быть только заказчиком IT и начала сама создавать и менять рабочую среду.\n\n"
        "Это не лекция про нейросети. Три живых кейса: оценка 360, бонусные письма и CV-сервис для тендеров. "
        "И что это меняет в роли человека, не только в технологиях.\n\n"
        "Переход: сначала — кто говорит и из какой роли.",
    )


# ---------------------------------------------------------------------------
# 02 about
# ---------------------------------------------------------------------------
def slide_about(prs):
    s = blank(prs)
    top = section(s, "Обо мне", "20 лет в HR.\nТеперь я создаю IT-решения", "02")
    facts = [
        ("01", "20+ лет", "в HR и управлении людьми в IT."),
        ("02", "HRD → COO", "От HR-функции к управлению бизнес-процессами."),
        ("03", "КРИТ", "IT-консалтинг, около 200 сотрудников."),
        ("04", "AI и бэк-офис", "Сама создаю процессы, а не только заказываю их."),
    ]
    y0 = top + Inches(0.08)
    for i, (num, title, text) in enumerate(facts):
        y = y0 + i * Inches(1.12)
        add_text(s, Inches(0.55), y, Inches(0.7), Inches(0.36), num, T_KICKER, ACCENT, pad=False)
        add_text(s, Inches(1.35), y, Inches(7.6), Inches(0.38), title, T_SUB, INK, font=HEAD, pad=False)
        add_text(s, Inches(1.35), y + Inches(0.40), Inches(7.6), Inches(0.36), text, T_BODY, MUTED, pad=False)
    place_image(s, SPEAKER, Inches(9.55), Inches(1.35), Inches(3.22), Inches(3.85), valign="top")
    add_text(
        s, Inches(9.55), Inches(5.32), Inches(3.22), Inches(1.05),
        "Раньше я была внутренним заказчиком IT. Сегодня сама создаю и меняю процессы.",
        T_CAPTION, INK, spacing=1.12, pad=False,
    )
    qr = s.shapes.add_picture(QR, Inches(9.55), Inches(6.42), height=Inches(0.58))
    qr.name = "QR Telegram"
    add_text(s, Inches(10.28), Inches(6.46), Inches(2.5), Inches(0.50), "Елена Никишева\nt.me/NikishevaTeam", T_CAPTION, INK, pad=False)
    notes(
        s, "02",
        "1:10. Четыре факта, без биографии.\n\n"
        "Двадцать лет я была внутри HR. Потом стала операционным директором. КРИТ — IT-консалтинг, около двухсот человек. "
        "И именно эта роль поменяла вопрос: не «какую систему купить», а «как устроена работа».\n\n"
        "Главный тезис слайда: раньше я была внутренним заказчиком IT. Сегодня сама собираю и меняю процессы. "
        "Не потому что стала программистом. Потому что AI и Cursor сняли часть барьера между идеей и первой рабочей версией.\n\n"
        "Переход: почему прежний путь автоматизации нас перестал устраивать.",
    )


# ---------------------------------------------------------------------------
# 03 old automation
# ---------------------------------------------------------------------------
def slide_old_path(prs):
    s = blank(prs)
    top = section(s, "Почему так", "От идеи до результата:\nслишком длинный путь", "03")
    add_text(s, Inches(0.55), top, Inches(12.2), Inches(0.28), "Как было", T_KICKER, MUTED, pad=False)
    old = ["Идея", "ТЗ", "Подрядчик", "Бюджет", "Разработка", "Доработки", "Внедрение"]
    y = top + Inches(0.42)
    add_box(s, Inches(0.55), y + Inches(0.16), Inches(12.22), Pt(2), LINE)
    for i, label in enumerate(old):
        x = Inches(0.55) + i * Inches(1.76)
        oval(s, x, y, Inches(0.36), Inches(0.36), MUTED)
        add_text(s, x, y, Inches(0.36), Inches(0.36), str(i + 1), T_CAPTION, WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=False)
        add_text(s, x - Inches(0.55), y + Inches(0.48), Inches(1.5), Inches(0.40), label, T_LABEL, INK, align=PP_ALIGN.CENTER, pad=False)
    add_text(
        s, Inches(0.55), y + Inches(1.05), Inches(12.2), Inches(0.36),
        "Зависимость от разработчиков  ·  Длинный цикл  ·  Затраты на подрядчика  ·  Чужая логика",
        T_BODY, MUTED, pad=False,
    )
    add_text(s, Inches(0.55), y + Inches(1.62), Inches(12.2), Inches(0.28), "Как стало возможным", T_KICKER, ACCENT, pad=False)
    new = ["Идея", "Прототип с AI", "Проверка", "Рабочий процесс"]
    nn = len(new)
    ngap = Inches(0.22)
    nw = (Inches(12.22) - (nn - 1) * ngap) / nn
    by = y + Inches(2.02)
    bh = Inches(1.35)
    for i, label in enumerate(new):
        x = Inches(0.55) + i * (nw + ngap)
        fill = ACCENT if i == 3 else BLUE_SOFT
        color = WHITE if i == 3 else INK
        add_box(s, x, by, nw, bh, fill)
        add_text(s, x, by + Inches(0.22), nw, Inches(0.28), f"{i + 1:02}", T_KICKER, WHITE if i == 3 else ACCENT, align=PP_ALIGN.CENTER, pad=False)
        add_text(s, x + Inches(0.1), by + Inches(0.58), nw - Inches(0.2), Inches(0.55), label, T_LEAD, color, align=PP_ALIGN.CENTER, pad=False)
        if i < nn - 1:
            arrow(s, x + nw + Inches(0.01), by + Inches(0.58), Inches(0.20), Inches(0.16), ACCENT)
    add_text(
        s, Inches(0.55), Inches(6.58), Inches(12.22), Inches(0.42),
        "Сроки и бюджеты сравнивать не будем. Меняется не смета — меняется, кто вправе сделать первый шаг.",
        T_CAPTION, MUTED, pad=False,
    )
    notes(
        s, "03",
        "1:20. Не называйте выдуманные сроки и экономию. Сравнения «было за полгода, стало за неделю» здесь нет, потому что так считать честно нельзя.\n\n"
        "Пройдите верхнюю цепочку пальцем: идея, ТЗ, подрядчик, бюджет, разработка, доработки, внедрение. "
        "Знакомо любому HRD. Четыре следствия: ждём разработчиков, цикл длинный, платим подрядчику, и на выходе часто чужая логика.\n\n"
        "Нижняя цепочка — не обещание, что всё мгновенно. Это другой вход: идею можно сразу пробовать. AI помогает собрать первую версию, мы проверяем, процесс идёт в работу.\n\n"
        "Тезис: главное изменение не в кнопке, а в том, кто может сделать первый шаг.\n\n"
        "Переход: первый кейс, где чужая логика нас остановила — оценка 360.",
    )


# ---------------------------------------------------------------------------
# 04 case 360
# ---------------------------------------------------------------------------
def slide_360_case(prs):
    s = blank(prs)
    top = section(s, "Кейс 1  ·  Оценка 360", "Готовое решение не подошло.\nМы сделали своё", "04")
    add_text(s, Inches(0.55), top, Inches(5.7), Inches(0.28), "Чужая логика", T_KICKER, MUTED, pad=False)
    add_text(
        s, Inches(0.55), top + Inches(0.36), Inches(5.7), Inches(1.05),
        "Готовое решение\nподрядчика",
        T_SUB, INK, font=HEAD, pad=False, spacing=1.02,
    )
    add_lines(
        s, Inches(0.55), top + Inches(1.55), Inches(5.5), Inches(2.4),
        [
            "Шаблон «как принято в продукте».",
            "HR подстраивается под систему.",
            "Нужная логика не совпала.",
            "Каждое изменение — через подрядчика.",
        ],
        T_BODY, MUTED, space=12, spacing=1.08,
    )
    add_box(s, Inches(6.70), top, Inches(6.08), Inches(5.18), BLUE_SOFT)
    add_text(s, Inches(6.98), top + Inches(0.22), Inches(5.55), Inches(0.28), "Своя логика", T_KICKER, ACCENT, pad=False)
    add_text(
        s, Inches(6.98), top + Inches(0.56), Inches(5.55), Inches(1.05),
        "Процесс под\nIT-консалтинг",
        T_SUB, INK, font=HEAD, pad=False, spacing=1.02,
    )
    add_lines(
        s, Inches(6.98), top + Inches(1.72), Inches(5.5), Inches(1.35),
        [
            "Оценщики из реальных проектов.",
            "Задачи, статусы и дашборды в Bitrix24.",
            "Правила процесса задаём сами.",
        ],
        T_BODY, INK, space=10, spacing=1.08,
    )
    add_text(s, Inches(6.98), top + Inches(3.22), Inches(5.5), Inches(0.78), "2 месяца", T_KPI, ACCENT, font=HEAD, pad=False)
    add_text(
        s, Inches(6.98), top + Inches(4.05), Inches(5.5), Inches(0.85),
        "вместо года, который планировали.\nСрок пересборки процесса, не экономия в деньгах.",
        T_CAPTION, MUTED, pad=False, spacing=1.12,
    )
    notes(
        s, "04",
        "1:20. История, не отчёт подрядчику.\n\n"
        "Нам нужна была оценка 360 с учётом специфики IT-консалтинга: проекты, роли, реальные пересечения людей. "
        "Готовое решение не совпало с этой логикой. Подстраивать свой процесс под чужой продукт мы не стали.\n\n"
        "Собрали сами в рабочей среде, в которой уже живёт компания, — в Bitrix24. "
        "То, что первоначально планировали на год, пересобрали за два месяца. Произнесите явно: это срок пересборки процесса, не подтверждённая экономия в деньгах.\n\n"
        "И сразу оговорка: 360 — не автономный AI-агент. Это пересобранный HR-процесс, который мы смогли сделать сами.\n\n"
        "Переход: как устроена автоматизация — от данных проектов до дашбордов, которые собираем сами.",
    )


# ---------------------------------------------------------------------------
# 05 360 system
# ---------------------------------------------------------------------------
def slide_360_system(prs):
    s = blank(prs)
    top = section(s, "Кейс 1  ·  Оценка 360", "От данных проектов\nдо решений, которые видим сами", "05")
    add_text(s, Inches(0.55), top, Inches(12.2), Inches(0.26), "Работает сейчас", T_KICKER, ACCENT, pad=False)
    steps = [
        ("01", "УПЗ", "Кто с кем\nработал на проектах"),
        ("02", "Назначение", "Список оценщиков\nна эту волну"),
        ("03", "Задачи", "Напоминания\nи статусы в Bitrix24"),
        ("04", "Анкеты", "Сбор обратной связи\nв одном контуре"),
        ("05", "Аналитика", "Дашборды, которые\nсобираем сами"),
    ]
    n = len(steps)
    gap = Inches(0.18)
    cw = (Inches(12.22) - (n - 1) * gap) / n
    y = top + Inches(0.38)
    ch = Inches(2.28)
    for i, (num, title, text) in enumerate(steps):
        x = Inches(0.55) + i * (cw + gap)
        fill = ACCENT if i == 4 else BLUE_SOFT
        tcol = WHITE if i == 4 else INK
        scol = WHITE if i == 4 else MUTED
        ncol = WHITE if i == 4 else ACCENT
        add_box(s, x, y, cw, ch, fill)
        add_text(s, x + Inches(0.12), y + Inches(0.14), cw - Inches(0.24), Inches(0.28), num, T_KICKER, ncol, pad=False)
        add_text(s, x + Inches(0.12), y + Inches(0.48), cw - Inches(0.24), Inches(0.52), title, T_LEAD, tcol, pad=False)
        add_text(s, x + Inches(0.12), y + Inches(1.12), cw - Inches(0.24), Inches(0.95), text, T_CAPTION, scol, pad=False, spacing=1.12)
        if i < n - 1:
            arrow(s, x + cw + Inches(0.01), y + Inches(1.02), Inches(0.16), Inches(0.14), ACCENT)
    add_text(
        s, Inches(0.55), y + ch + Inches(0.18), Inches(12.2), Inches(0.26),
        "Подключаем дальше — ещё не готово", T_KICKER, MUTED, pad=False,
    )
    add_text(
        s, Inches(0.55), y + ch + Inches(0.48), Inches(12.2), Inches(0.42),
        "Утилизация — дорабатываем правила и источники.  ·  Рекомендации по живой коммуникации — почта и Teams.",
        T_BODY, MUTED, pad=False,
    )
    band_t = y + ch + Inches(1.12)
    add_box(s, Inches(0.55), band_t, Inches(12.22), Inches(0.82), NAVY)
    add_text(
        s, Inches(0.78), band_t, Inches(11.76), Inches(0.82),
        "Нужный срез и логику процесса собираем сами. Добавить дашборд — не заявка подрядчику.",
        T_LEAD, WHITE, anchor=MSO_ANCHOR.MIDDLE, pad=False,
    )
    notes(
        s, "05",
        "1:40. Это схема процесса, не техническая архитектура. Не говорите, что все блоки уже связаны API.\n\n"
        "Пройдите цепочку слева направо. УПЗ даёт, кто с кем работал на проектах. Дальше назначение оценщиков, задачи со статусами, анкеты. "
        "Последний шаг — аналитика: дашборды собираем сами под нужный срез, а не ждём отчёт от вендора.\n\n"
        "Утилизацию на слайде не ставьте в один ряд с работающим контуром. Колонка в отчёте уже есть, но правила и источники мы ещё дорабатываем. "
        "Рекомендации по живой коммуникации — почта и Teams — следующий шаг. Это не сделанный контур.\n\n"
        "Тезис нижней полосы: нужный дашборд и логику процесса мы задаём сами.\n\n"
        "Переход: как это выглядит в рабочей среде — не процентами сдачи, а управляемостью.",
    )


# ---------------------------------------------------------------------------
# 06 360 results
# ---------------------------------------------------------------------------
def slide_360_results(prs):
    s = blank(prs)
    top = section(s, "Кейс 1  ·  Оценка 360", "Ход оценки виден.\nЛогику меняем сами", "06")
    points = [
        ("01", "Ход волны", "Календарь, статусы и задачи живут в Bitrix24, а не в отдельной таблице."),
        ("02", "Свои дашборды", "Свод, срезы и отчёт для бонуса и ИПР собираем под нужное решение."),
        ("03", "Своя логика", "Правила процесса меняем сами. Это не заявка подрядчику и не новый IT-проект."),
    ]
    y = top
    for num, title, text in points:
        add_text(s, Inches(0.55), y, Inches(0.55), Inches(0.32), num, T_KICKER, ACCENT, pad=False)
        add_text(s, Inches(1.18), y, Inches(4.85), Inches(0.34), title, T_SUB, INK, font=HEAD, pad=False)
        add_text(s, Inches(1.18), y + Inches(0.38), Inches(4.85), Inches(0.72), text, T_BODY, MUTED, pad=False, spacing=1.08)
        y += Inches(1.22)
    add_text(
        s, Inches(0.55), Inches(6.28), Inches(5.5), Inches(0.68),
        "Контур оставляем. Дальше делаем честнее то, что стоит рядом: утилизацию и список оценивающих по живой работе.",
        T_CAPTION, MUTED, pad=False, spacing=1.12,
    )
    shot_l, shot_t = Inches(6.55), top
    shot_w, shot_h = Inches(6.22), Inches(4.48)
    add_box(s, shot_l, shot_t, shot_w, shot_h, SOFT)
    pic = place_image(
        s, SHOT_360_DASH,
        shot_l + Inches(0.08), shot_t + Inches(0.08),
        shot_w - Inches(0.16), shot_h - Inches(0.16),
        mode="contain", valign="top",
    )
    pic.name = "Сводный отчёт оценки 360"
    add_text(
        s, Inches(6.55), Inches(6.28), Inches(6.22), Inches(0.42),
        "Сводный отчёт 360 в Bitrix24. Имена и баллы скрыты.",
        T_CAPTION, MUTED, pad=False,
    )
    notes(
        s, "06",
        "1:40. Не читайте проценты сдачи и не сравнивайте волны цифрами. На слайде этого нет — и в зале это не нужно.\n\n"
        "Покажите экран: ход оценки виден в той же среде, где работает компания. Дашборды собираем сами под бонус, ИПР и разговор с руководителем. "
        "Если нужна другая логика или другой срез — меняем сами.\n\n"
        "Контур 360 оставляем. Следующий шаг не «изобрести оценку заново», а сделать честнее то, что стоит рядом: утилизацию и список оценивающих по живой коммуникации. "
        "Следующая волна — июнь 2027.\n\n"
        "Переход: второй кейс — совсем другого масштаба. Не год и не два месяца. Полдня.",
    )


# ---------------------------------------------------------------------------
# 07 bonus letters
# ---------------------------------------------------------------------------
def slide_bonus_case(prs):
    s = blank(prs)
    add_box(s, 0, 0, Inches(4.85), H, ACCENT)
    s.shapes.add_picture(LOGO_WHITE, Inches(0.42), Inches(0.28), height=Inches(0.42))
    add_text(s, Inches(0.42), Inches(1.15), Inches(4.05), Inches(0.32), "Кейс 2  ·  Бонусные письма", T_KICKER, WHITE, pad=False)
    add_text(s, Inches(0.38), Inches(1.85), Inches(4.15), Inches(1.55), "½", 56, WHITE, font=HEAD, pad=False)
    add_text(
        s, Inches(0.42), Inches(3.55), Inches(4.05), Inches(1.05),
        "рабочего дня\nна первую версию",
        T_LEAD, WHITE, pad=False, spacing=1.12,
    )
    add_text(
        s, Inches(0.42), Inches(5.22), Inches(4.05), Inches(1.35),
        "Не полный цикл тестирования и внедрения. Первая рабочая версия, которую можно было открыть и пройти.",
        T_CAPTION, WHITE, pad=False, spacing=1.12,
    )
    add_text(s, Inches(5.28), Inches(0.22), Inches(6.4), Inches(0.24), "Личная история", T_KICKER, ACCENT, pad=False)
    add_text(
        s, Inches(5.28), Inches(0.48), Inches(7.5), Inches(0.85),
        "Полдня, чтобы создать\nсобственный процесс",
        T_TITLE, INK, font=HEAD, pad=False, spacing=0.98,
    )
    add_text(s, Inches(12.55), Inches(0.22), Inches(0.4), Inches(0.24), "07", T_CAPTION, MUTED, align=PP_ALIGN.RIGHT, pad=False)
    add_text(s, Inches(5.28), Inches(1.62), Inches(3.55), Inches(0.36), "Было", T_SUB, MUTED, font=HEAD, pad=False)
    add_lines(
        s, Inches(5.28), Inches(2.15), Inches(3.4), Inches(4.2),
        [
            "Excel-файлы.",
            "Разные модули.",
            "Отдельные операции.",
            "Связка действий\nдержится на памяти.",
        ],
        T_BODY, INK, space=14, spacing=1.08,
    )
    add_box(s, Inches(8.95), Inches(1.55), Inches(3.85), Inches(5.15), BLUE_SOFT)
    add_text(s, Inches(9.16), Inches(1.74), Inches(3.45), Inches(0.36), "Стало", T_SUB, ACCENT, font=HEAD, pad=False)
    add_lines(
        s, Inches(9.16), Inches(2.22), Inches(3.45), Inches(4.2),
        [
            "Свой смарт-процесс\nв Bitrix24.",
            "Постановка целей.",
            "Проверка руководителями.",
            "Согласование.",
            "Единый маршрут.",
        ],
        T_BODY, INK, space=12, spacing=1.08,
    )
    notes(
        s, "07",
        "1:20. Это личная история, говорите от первого лица.\n\n"
        "Бонусные письма жили в Excel и разъезжались по модулям. Чтобы провести цикл, нужно было помнить, что за чем следует. "
        "Неудобно не потому что люди ленивые — потому что процесс не был собран в одном месте.\n\n"
        "Я собрала первую рабочую версию своего смарт-процесса примерно за половину рабочего дня: цели, проверка руководителем, согласование, один маршрут. "
        "Сразу скажите: это не внедрение «под ключ» и не полный цикл тестирования. Это момент, когда процесс впервые можно пройти руками.\n\n"
        "Тезис: изменение небольшого процесса больше не обязано быть IT-проектом.\n\n"
        "Переход: что изменилось после того, как процесс стал своим.",
    )


# ---------------------------------------------------------------------------
# 08 bonus change
# ---------------------------------------------------------------------------
def slide_bonus_change(prs):
    s = blank(prs)
    top = section(s, "Кейс 2  ·  Бонусные письма", "Изменение процесса —\nтеперь часть моей работы", "08")
    steps = [
        ("01", "Проблема"),
        ("02", "Идея"),
        ("03", "С AI"),
        ("04", "Проверка"),
        ("05", "Запуск"),
        ("06", "Улучшение"),
    ]
    n = len(steps)
    gap = Inches(0.16)
    cw = (Inches(12.22) - (n - 1) * gap) / n
    y = top
    ch = Inches(1.28)
    for i, (num, label) in enumerate(steps):
        x = Inches(0.55) + i * (cw + gap)
        live = i in (2, 5)
        add_box(s, x, y, cw, ch, ACCENT if live else BLUE_SOFT)
        color = WHITE if live else INK
        ncolor = WHITE if live else ACCENT
        add_text(s, x + Inches(0.08), y + Inches(0.10), cw - Inches(0.16), Inches(0.26), num, T_KICKER, ncolor, pad=False)
        add_text(s, x + Inches(0.08), y + Inches(0.48), cw - Inches(0.16), Inches(0.58), label, T_BODY, color, pad=False)
        if i < n - 1:
            arrow(s, x + cw + Inches(0.01), y + Inches(0.56), Inches(0.14), Inches(0.14), ACCENT)
    add_text(
        s, Inches(0.55), y + ch + Inches(0.28), Inches(5.55), Inches(1.15),
        "Правила, поля и маршрут согласования меняем сами. Небольшой сдвиг больше не обязан быть IT-проектом.",
        T_BODY, INK, pad=False, spacing=1.08,
    )
    shot_top = y + ch + Inches(0.18)
    pic = place_image(
        s, SHOT_BONUS,
        Inches(6.45), shot_top,
        Inches(6.32), Inches(6.72) - shot_top,
        mode="contain", valign="top",
    )
    add_text(
        s, Inches(6.45), Inches(6.72), Inches(6.32), Inches(0.24),
        "Страница бонусного письма в Bitrix24",
        T_CAPTION, MUTED, pad=False,
    )
    pic.name = "Бонусное письмо — цели и маршрут"
    notes(
        s, "08",
        "1:10. Не пересказывайте поля на скрине. Покажите маршрут: цели, согласование, подпись, результаты — один процесс.\n\n"
        "Главное: когда процесс свой, изменение правил — часть моей работы, а не заявка в IT. "
        "Жизненный цикл короткий: увидел проблему, придумал, с помощью AI собрал, проверил, запустил, улучшил.\n\n"
        "Скрин настоящий, из рабочего контура бонусных писем. Имена на учебном экране скрыты. "
        "Не выдавайте первую версию за законченную систему: процесс живой, мы его продолжаем менять.\n\n"
        "Переход: третий кейс — там ставка уже не удобство HR, а деньги компании на тендерах.",
    )


# ---------------------------------------------------------------------------
# 09 CV problems
# ---------------------------------------------------------------------------
def slide_cv_problems(prs):
    s = blank(prs)
    top = section(s, "Кейс 3  ·  CV-сервис", "Резюме должны работать на бизнес,\nа не лежать в папках", "09")
    add_text(
        s, Inches(0.55), top, Inches(12.2), Inches(0.42),
        "КРИТ — IT-консалтинг. Мы ходим в тендеры. Актуальность опыта — не HR-аккуратность, а готовность команды к сделке.",
        T_BODY, INK, pad=False,
    )
    items = [
        ("01", "Актуальность CV", "Не видно, кто обновил резюме и когда. Резюме живут сами по себе."),
        ("02", "Потеря проектов", "Человек вышел из проекта — и через полгода опыт из резюме уже не вспомнить."),
        ("03", "Тендер не ждёт", "Нужный опыт надо найти и успеть подготовить людей под конкретную закупку."),
    ]
    y = top + Inches(0.55)
    widths = [Inches(9.2), Inches(10.55), Inches(12.22)]
    fills = [SOFT, BLUE_SOFT, ACCENT]
    colors = [INK, INK, WHITE]
    muted_c = [MUTED, MUTED, WHITE]
    for i, ((num, title, text), w, fill, c, mc) in enumerate(zip(items, widths, fills, colors, muted_c)):
        add_box(s, Inches(0.55), y, w, Inches(1.22), fill)
        add_text(s, Inches(0.75), y + Inches(0.16), Inches(0.7), Inches(0.34), num, T_KICKER, ACCENT if i < 2 else WHITE, pad=False)
        add_text(s, Inches(1.55), y + Inches(0.16), w - Inches(1.3), Inches(0.34), title, T_SUB, c, font=HEAD, pad=False)
        add_text(s, Inches(1.55), y + Inches(0.56), w - Inches(1.3), Inches(0.50), text, T_BODY, mc, pad=False)
        y += Inches(1.34)
    add_text(
        s, Inches(0.55), Inches(6.58), Inches(12.22), Inches(0.40),
        "Три проблемы закрываем в одном месте: Bitrix24 и цифровые сотрудники.",
        T_CAPTION, MUTED, pad=False,
    )
    notes(
        s, "09",
        "1:10. Соберите три проблемы в одну историю, не читайте карточки.\n\n"
        "Резюме устаревают незаметно. Проекты забываются, как только человек перешёл на следующий. "
        "А когда появляется тендер, нужный опыт надо найти быстро — иначе команда собирается по памяти и по папкам.\n\n"
        "Для IT-консалтинга это не косметика HR. Это готовность выйти в закупку с релевантным опытом.\n\n"
        "Тезис: резюме должны работать на бизнес, а не лежать в папках.\n\n"
        "Переход: как это устроено, двумя ветками — обновление CV и подготовка к тендеру.",
    )


# ---------------------------------------------------------------------------
# 10 CV how it works
# ---------------------------------------------------------------------------
def slide_cv_flow(prs):
    s = blank(prs)
    top = section(s, "Кейс 3  ·  CV-сервис", "От данных о проекте до готового CV", "10")
    points = [
        ("01", "Задача с данными проекта", "Вход или выход из проекта даёт задачу уже с названием, заказчиком, периодом и ролью. Сотрудник проверяет, HR согласует."),
        ("02", "Поиск опыта под тендер", "Помощник разбирает требования и ищет опыт в CV-сервисе. Список проверяют РП или ДД, затем HR готовит резюме."),
        ("03", "Отчёт «Резюме»", "Реестр уже в Bitrix24: кто в списке, должность, ссылка. Даты файла с SharePoint ещё подключаем."),
    ]
    y = top
    for num, title, text in points:
        add_text(s, Inches(0.55), y, Inches(0.55), Inches(0.32), num, T_KICKER, ACCENT, pad=False)
        add_text(s, Inches(1.18), y, Inches(4.85), Inches(0.34), title, T_SUB, INK, font=HEAD, pad=False)
        add_text(s, Inches(1.18), y + Inches(0.38), Inches(4.85), Inches(0.78), text, T_BODY, MUTED, pad=False, spacing=1.08)
        y += Inches(1.22)
    add_text(
        s, Inches(0.55), Inches(6.28), Inches(5.5), Inches(0.68),
        "Цифровые помощники берут подготовку. Человек оставляет проверку и ответственность.",
        T_CAPTION, MUTED, pad=False, spacing=1.12,
    )
    shot_l, shot_t = Inches(6.55), top
    shot_w, shot_h = Inches(6.22), Inches(4.48)
    add_box(s, shot_l, shot_t, shot_w, shot_h, SOFT)
    pic = place_image(
        s, SHOT_CV_REPORT,
        shot_l + Inches(0.08), shot_t + Inches(0.08),
        shot_w - Inches(0.16), shot_h - Inches(0.16),
        mode="contain", valign="top",
    )
    pic.name = "Отчёт по резюме в Bitrix24"
    add_text(
        s, Inches(6.55), Inches(6.28), Inches(6.22), Inches(0.42),
        "Отчёт по резюме в Bitrix24. Имена скрыты.",
        T_CAPTION, MUTED, pad=False,
    )
    notes(
        s, "10",
        "1:40. Две ветки, не архитектура систем. Экран справа — живой отчёт по резюме в Bitrix24, как свод 360 на шестом слайде.\n\n"
        "Первая: человек входит в проект или выходит — из управления проектами приходит сигнал — в задачу обновления резюме уже подставлены название, заказчик, период, роль. "
        "Сотрудник не сочиняет проект заново, он проверяет. HR согласовывает.\n\n"
        "Вторая: решили идти в тендер — помощник разбирает, какие навыки нужны — ищет опыт в CV-сервисе — отдаёт список РП или ДД. "
        "Если просят найти ещё, поиск повторяется. Когда «ок», задача падает на HR: кого готовить и в какую роль.\n\n"
        "На экране — реестр: кто в списке, должность, ссылка на файл. Даты обновления с SharePoint ещё не подгружаются, это не скрывать и не дорисовывать. "
        "Число сотрудников на скрине не делайте KPI и не называйте экономию: измеренного эффекта по контуру нет.\n\n"
        "Переход: что общего у трёх разных кейсов.",
    )


# ---------------------------------------------------------------------------
# 11 what unites
# ---------------------------------------------------------------------------
def slide_unites(prs):
    s = blank(prs, ACCENT)
    add_text(s, Inches(0.62), Inches(0.38), Inches(10.5), Inches(0.28), "Три кейса — один сдвиг", T_KICKER, WHITE, pad=False)
    add_text(s, Inches(11.9), Inches(0.38), Inches(0.9), Inches(0.28), "11", T_KICKER, WHITE, align=PP_ALIGN.RIGHT, pad=False)
    add_text(
        s, Inches(0.62), Inches(0.78), Inches(12.1), Inches(1.15),
        "Мы перестали ждать.\nМы начали создавать.",
        T_DISPLAY, WHITE, font=HEAD, pad=False, spacing=0.96,
    )
    add_box(s, Inches(0.62), Inches(2.22), Inches(3.85), Inches(2.72), WHITE)
    add_text(s, Inches(0.84), Inches(2.40), Inches(3.45), Inches(0.28), "01", T_KICKER, ACCENT, pad=False)
    add_text(s, Inches(0.84), Inches(2.76), Inches(3.45), Inches(0.70), "Делаем сами", T_SUB, INK, font=HEAD, pad=False, spacing=1.02)
    add_text(s, Inches(0.84), Inches(3.62), Inches(3.45), Inches(1.05), "Можем реализовать собственную логику процесса.", T_BODY, MUTED, pad=False, spacing=1.08)
    add_box(s, Inches(4.65), Inches(2.22), Inches(3.85), Inches(2.72), NAVY)
    add_text(s, Inches(4.87), Inches(2.40), Inches(3.45), Inches(0.28), "02", T_KICKER, ACCENT, pad=False)
    add_text(s, Inches(4.87), Inches(2.76), Inches(3.45), Inches(0.50), "Скорость", T_SUB, WHITE, font=HEAD, pad=False)
    add_text(s, Inches(4.87), Inches(3.62), Inches(3.45), Inches(1.05), "Быстрее проверяем идеи и собираем первые версии.", T_BODY, WHITE, pad=False, spacing=1.08)
    outline(s, Inches(8.68), Inches(2.22), Inches(3.85), Inches(2.72), WHITE, 1.5)
    add_text(s, Inches(8.90), Inches(2.40), Inches(3.45), Inches(0.28), "03", T_KICKER, WHITE, pad=False)
    add_text(s, Inches(8.90), Inches(2.76), Inches(3.45), Inches(0.70), "Новая\nэкспертиза", T_SUB, WHITE, font=HEAD, pad=False, spacing=1.02)
    add_text(s, Inches(8.90), Inches(3.62), Inches(3.45), Inches(1.05), "HR начинает понимать процессы, данные и возможности разработки.", T_BODY, WHITE, pad=False, spacing=1.08)
    add_text(
        s, Inches(0.62), Inches(5.22), Inches(12.1), Inches(1.45),
        "HR не становится профессиональным программистом.\nЕго возможности больше не ограничены умением написать ТЗ.",
        T_LEAD, WHITE, pad=False, spacing=1.15,
    )
    notes(
        s, "11",
        "1:00. Короткий переходный слайд. Не перечисляйте кейсы заново.\n\n"
        "Три изменения: самостоятельность — можем собрать свою логику. Скорость — быстрее доходим до первой версии. "
        "Новая экспертиза — HR начинает понимать процессы, данные и то, что вообще можно сделать в разработке.\n\n"
        "И обязательная граница: мы не готовим из HR программистов. Мы снимаем потолок, который раньше стоял на умении написать техническое задание.\n\n"
        "Переход: если транзакции уходят AI, что тогда остаётся человеческой работой в HR?",
    )


# ---------------------------------------------------------------------------
# 12 future HR
# ---------------------------------------------------------------------------
def slide_future(prs):
    s = blank(prs)
    top = section(
        s, "В логике исследований Джоша Берсина",
        "Если транзакции уходят AI,\nчто остаётся человеку?",
        "12",
    )
    add_text(s, Inches(0.55), top, Inches(5.7), Inches(0.32), "Транзакционная работа", T_KICKER, MUTED, pad=False)
    add_text(s, Inches(7.05), top, Inches(5.7), Inches(0.32), "Новая ценность HR", T_KICKER, ACCENT, pad=False)
    add_text(
        s, Inches(0.55), top + Inches(0.48), Inches(5.5), Inches(2.55),
        "Рутину всё чаще берёт автоматизация.\n\n"
        "Формы, статусы, сбор данных,\nнапоминания, первичная подготовка.\n\n"
        "Это сдвиг нагрузки, а не прогноз,\nчто HR исчезнет.",
        T_BODY, MUTED, pad=False, spacing=1.12,
    )
    add_box(s, Inches(6.85), top + Inches(0.40), Inches(5.92), Inches(2.85), BLUE_SOFT)
    add_text(
        s, Inches(7.08), top + Inches(0.58), Inches(5.5), Inches(2.50),
        "Проектирование процессов.\n"
        "Организационные изменения.\n"
        "Сложные решения и ответственность.\n"
        "Данные, AI и устройство работы.",
        T_LEAD, INK, pad=False, spacing=1.22,
    )
    add_text(
        s, Inches(0.55), Inches(5.55), Inches(12.22), Inches(1.15),
        "Будущее HR — не в количестве автоматизированных операций.\n"
        "А в способности иначе организовать работу людей и технологий.",
        T_SUB, INK, font=HEAD, pad=False, spacing=1.12,
    )
    notes(
        s, "12",
        "1:30. Не цитируйте Берсина и не называйте проценты автоматизации HR — это чужие оценки, не наши факты.\n\n"
        "Скажите так: в исследованиях Джоша Берсина тот же вопрос, который мы видим у себя. "
        "Если транзакционную работу всё больше забирает AI, ценность HR смещается к проектированию процессов, к изменениям в организации и к решениям, которые нельзя отдать машине.\n\n"
        "Слева — то, что уходит в автоматизацию. Справа — то, что остаётся человеческим и даже возрастает: данные, AI, устройство работы.\n\n"
        "Тезис: будущее не в том, сколько операций мы закрыли роботом. А в том, умеем ли иначе собрать работу людей и технологий.\n\n"
        "Переход: этот сдвиг нельзя оставить привилегией одного руководителя. Поэтому в КРИТ запущен AI Challenge.",
    )


# ---------------------------------------------------------------------------
# 13 AI Challenge
# ---------------------------------------------------------------------------
def slide_challenge(prs):
    s = blank(prs)
    top = section(s, "AI Challenge", "Следующий шаг —\nчтобы разработчиками стали все", "13")
    add_text(
        s, Inches(0.55), top, Inches(12.2), Inches(0.42),
        "Не соревнование ради ботов. Управленческая программа: научить бэк-офис менять свою работу.",
        T_BODY, INK, pad=False,
    )
    cy = top + Inches(0.52)
    ch = Inches(0.58)
    gap_y = Inches(0.08)
    block_h = 6 * ch + 5 * gap_y
    add_box(s, Inches(0.55), cy, Inches(4.35), block_h, NAVY)
    add_text(s, Inches(0.78), cy + Inches(0.22), Inches(3.95), Inches(0.28), "От одного к команде", T_KICKER, ACCENT, pad=False)
    add_text(
        s, Inches(0.78), cy + Inches(0.62), Inches(3.95), Inches(1.55),
        "Сначала процессы\nсобирал руководитель.\nТеперь это возможность\nвсей команды.",
        T_LEAD, WHITE, pad=False, spacing=1.12,
    )
    add_text(
        s, Inches(0.78), cy + Inches(2.40), Inches(3.95), Inches(1.35),
        "Найти процесс.\nСобрать прототип.\nПоказать результат.\nЛучшее — в работу.",
        T_BODY, WHITE, pad=False, spacing=1.18,
    )
    cycle = [
        ("01", "Нашёл проблему"),
        ("02", "Предложил идею"),
        ("03", "Создал прототип"),
        ("04", "Показал результат"),
        ("05", "Внедрил"),
        ("06", "Продолжил улучшать"),
    ]
    for i, (num, label) in enumerate(cycle):
        y = cy + i * (ch + gap_y)
        add_text(s, Inches(5.28), y, Inches(0.7), ch, num, T_KICKER, ACCENT, anchor=MSO_ANCHOR.MIDDLE, pad=False)
        add_text(s, Inches(6.10), y, Inches(6.5), ch, label, T_LEAD, INK, anchor=MSO_ANCHOR.MIDDLE, pad=False)
        if i < 5:
            add_box(s, Inches(5.28), y + ch + Inches(0.02), Inches(7.5), Pt(1), LINE)
    add_text(
        s, Inches(0.55), cy + block_h + Inches(0.14), Inches(12.22), Inches(0.36),
        "Challenge продолжается. По каждому прототипу решение отдельное: внедрить, доработать или остановить.",
        T_CAPTION, MUTED, pad=False,
    )
    notes(
        s, "13",
        "1:40. Не превращайте слайд в отчёт о хакатоне.\n\n"
        "AI Challenge в КРИТ — внутренняя программа для бэк-офиса. Задача не наплодить ботов. "
        "Задача — научить людей видеть процесс, который можно изменить, собрать прототип и, если он работает, поставить его в живую работу.\n\n"
        "Пройдите цикл один раз: нашёл проблему, предложил идею, создал прототип, показал, внедрил, продолжил улучшать. "
        "Cursor и AI здесь инструменты, не цель.\n\n"
        "Произнесите статус прямо: программа живая. Мы не внедрили «все решения». По каждому прототипу решение отдельное — доработать, внедрить или остановить эксперимент.\n\n"
        "Тезис: следующий шаг — чтобы возможность создавать процесс была не у одного COO, а у команды.\n\n"
        "Переход: финальная мысль. Не список выводов — одна фраза, за которой можно остановиться.",
    )


# ---------------------------------------------------------------------------
# 14 finale
# ---------------------------------------------------------------------------
def slide_final(prs):
    s = blank(prs, NAVY)
    s.shapes.add_picture(LOGO_WHITE, Inches(0.55), Inches(0.32), height=Inches(0.44))
    add_text(s, Inches(11.4), Inches(0.40), Inches(1.4), Inches(0.28), "14", T_KICKER, ACCENT, align=PP_ALIGN.RIGHT, pad=False)
    add_text(
        s, Inches(0.62), Inches(1.15), Inches(12.1), Inches(1.85),
        "«Единственный способ делать\nдействительно большую работу —\nлюбить то, что ты делаешь».",
        T_TITLE, WHITE, font=HEAD, pad=False, spacing=1.10,
    )
    add_text(
        s, Inches(0.62), Inches(3.12), Inches(12.1), Inches(0.36),
        "Стив Джобс  ·  выступление в Стэнфорде, 2005",
        T_KICKER, ACCENT, pad=False,
    )
    add_box(s, Inches(0.62), Inches(3.70), Inches(1.45), Pt(2.5), ACCENT)
    add_text(
        s, Inches(0.62), Inches(4.00), Inches(12.1), Inches(1.25),
        "Мы начинали с автоматизации процессов.\nА теперь меняем представление о том, что можем делать сами.",
        T_SUB, WHITE, font=HEAD, pad=False, spacing=1.15,
    )
    s.shapes.add_picture(VERTIK, Inches(0.52), Inches(5.72), height=Inches(1.18))
    add_text(s, Inches(1.90), Inches(5.92), Inches(5.5), Inches(0.36), "Елена Никишева  ·  COO КРИТ", T_BODY, WHITE, pad=False)
    add_text(s, Inches(1.90), Inches(6.32), Inches(5.5), Inches(0.30), "t.me/NikishevaTeam", T_KICKER, ACCENT, pad=False)
    qr = s.shapes.add_picture(QR, Inches(11.55), Inches(5.72), height=Inches(1.18))
    qr.name = "QR Telegram"
    notes(
        s, "14",
        "1:00. Не зачитывайте цитату как лозунг компании. Это чужие слова — Джобс, Стэнфорд, 2005. Пауза.\n\n"
        "Затем ваша фраза, отдельно, не как продолжение цитаты: мы начинали с автоматизации процессов. "
        "А теперь меняем представление о том, что можем делать сами.\n\n"
        "Остановитесь. Имя, роль, Telegram — если кто-то захочет продолжить разговор. "
        "Не добавляйте третьего вывода и не открывайте следующий слайд: его нет.",
    )


def write_notes_md():
    order = [
        ("01", "Титульный", "0:40"),
        ("02", "Обо мне", "1:10"),
        ("03", "Почему прежняя автоматизация перестала устраивать", "1:20"),
        ("04", "Кейс 1. Оценка 360", "1:20"),
        ("05", "Оценка 360: что мы соединили", "1:40"),
        ("06", "Оценка 360: ход волны и своя логика", "1:40"),
        ("07", "Кейс 2. Бонусные письма", "1:20"),
        ("08", "Бонусные письма: что изменилось", "1:10"),
        ("09", "Кейс 3. CV-сервис для тендеров", "1:10"),
        ("10", "Как устроен CV-сервис", "1:40"),
        ("11", "Что объединяет все три кейса", "1:00"),
        ("12", "HR будущего", "1:30"),
        ("13", "AI Challenge", "1:40"),
        ("14", "Финал", "1:00"),
    ]
    lines = [
        "# Сценарий выступления: HR API, 1 декабря 2026",
        "",
        "**Тема.** HR сам себе разработчик: как мы создаём и внедряем ИИ-инструменты в КРИТ.",
        "**Спикер.** Елена Никишева, COO КРИТ. Спикер один.",
        "**Хронометраж.** 20 минут. В сумме ниже — 18:20, запас 1:40 на паузы.",
        "**Тон.** Уверенный, живой, практический. Без корпоративных клише и технического жаргона.",
        "",
        "Главный тезис, который нужно удержать до конца: раньше для автоматизации HR нужны были подрядчик, бюджет, ТЗ и месяцы ожидания. "
        "Сейчас HR может сам создавать решения, проверять гипотезы и менять рабочую среду. Меняется не только технология — меняется профессиональная роль.",
        "",
        "## Тайминг",
        "",
        "| Слайд | Тема | Время |",
        "| --- | --- | --- |",
    ]
    for num, title, t in order:
        lines.append(f"| {num} | {title} | {t} |")
    lines += ["", "---", ""]
    for num, title, t in order:
        lines += [
            f"## Слайд {num}. {title}",
            f"**Время:** {t}",
            "",
            NOTES[num],
            "",
            "---",
            "",
        ]
    OUT_NOTES.write_text("\n".join(lines), encoding="utf-8")


def set_theme_fonts(prs):
    """Тема как в рабочих PPTX КРИТ: заголовки Uncage, текст Arial."""
    rels = prs.slide_masters[0].part.rels
    theme = None
    for rel in rels.values():
        if "theme" in str(rel.reltype).lower():
            theme = rel.target_part
            break
    if theme is None:
        return
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    root = etree.fromstring(theme.blob)
    major = root.find(f".//{ns}majorFont/{ns}latin")
    minor = root.find(f".//{ns}minorFont/{ns}latin")
    if major is not None:
        major.set("typeface", HEAD)
    if minor is not None:
        minor.set("typeface", BODY)
    theme._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    set_theme_fonts(prs)
    prs.core_properties.title = "HR сам себе разработчик: как мы создаём и внедряем ИИ-инструменты в КРИТ"
    prs.core_properties.author = "Елена Никишева"
    prs.core_properties.subject = "HR API, 1 декабря 2026"
    prs.core_properties.category = "КРИТ"
    slide_cover(prs)
    slide_about(prs)
    slide_old_path(prs)
    slide_360_case(prs)
    slide_360_system(prs)
    slide_360_results(prs)
    slide_bonus_case(prs)
    slide_bonus_change(prs)
    slide_cv_problems(prs)
    slide_cv_flow(prs)
    slide_unites(prs)
    slide_future(prs)
    slide_challenge(prs)
    slide_final(prs)
    prs.save(OUT_PPTX)
    write_notes_md()
    print(OUT_PPTX)
    print(OUT_NOTES)
    return OUT_PPTX


if __name__ == "__main__":
    main()
