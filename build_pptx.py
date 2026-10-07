#!/usr/bin/env python3
"""Build the AI Challenge employee deck as a 16:9 PowerPoint."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from pptx.oxml import parse_xml
from lxml import etree

W = Inches(13.333)
H = Inches(7.5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1E, 0x1E, 0x1E)
MUTED = RGBColor(0x77, 0x77, 0x77)
ACCENT = RGBColor(0x38, 0x9B, 0xFF)
LINE = RGBColor(0xE0, 0xE3, 0xE8)
SOFT = RGBColor(0xF4, 0xF8, 0xFC)
HEAD = "Uncage"
BODY = "Arial"


def set_run(run, text, size, color, font=BODY, bold=False):
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = font
    run.font.bold = bold
    rPr = run._r.get_or_add_rPr()
    # keep East Asian / complex scripts on Arial so Cyrillic does not jump fonts
    for tag in ("ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", font)


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


def add_bar(slide, l, t, w, fill=ACCENT, h=Pt(3)):
    return add_box(slide, l, t, w, h, fill)


def tf_shape(slide, l, t, w, h):
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    return sh, tf


def para(tf, text, size, color, font=BODY, bold=False, align=PP_ALIGN.LEFT, space_after=0, first=False):
    p = tf.paragraphs[0] if first and tf.paragraphs[0].text == "" else tf.add_paragraph()
    if first:
        p = tf.paragraphs[0]
        p.clear()
    p.alignment = align
    p.space_after = Pt(space_after)
    run = p.add_run()
    set_run(run, text, size, color, font=font, bold=bold)
    return p


def add_text(slide, l, t, w, h, text, size, color, font=BODY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    sh, tf = tf_shape(slide, l, t, w, h)
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    para(tf, text, size, color, font=font, bold=bold, align=align, first=True)
    return sh


def card(slide, l, t, w, h, fill=WHITE, accent=True):
    add_box(slide, l, t, w, h, fill, LINE)
    if accent:
        add_bar(slide, l, t, w)
    return l, t, w, h


def header(slide, kicker, title, logo, num):
    add_box(slide, Inches(0.28), Inches(1.18), Pt(2.5), Inches(5.55), ACCENT)
    add_text(slide, Inches(0.55), Inches(0.22), Inches(10.4), Inches(0.32), kicker, 12, ACCENT)
    add_text(slide, Inches(0.55), Inches(0.48), Inches(10.4), Inches(0.55), title, 22, INK, font=HEAD)
    slide.shapes.add_picture(logo, Inches(11.85), Inches(0.28), height=Inches(0.42))
    add_text(slide, Inches(0.55), Inches(7.12), Inches(8), Inches(0.28), "AI Challenge · ОЦО и бэкофис", 10, ACCENT)
    add_text(slide, Inches(12.35), Inches(7.12), Inches(0.6), Inches(0.28), num, 10, MUTED, align=PP_ALIGN.RIGHT)


def logo_grey():
    return "assets/pptx/logo-grey.png"


def logo_white():
    return "assets/pptx/logo-white.png"


def cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    builder = s.shapes.build_freeform(Inches(10.55), Inches(0))
    builder.add_line_segments(
        [(W, Inches(0)), (W, H), (Inches(8.82), H)],
        close=True,
    )
    geo = builder.convert_to_shape()
    geo.shadow.inherit = False
    geo.fill.solid()
    geo.fill.fore_color.rgb = ACCENT
    geo.line.fill.background()
    s.shapes.add_picture(logo_white(), Inches(11.85), Inches(0.32), height=Inches(0.42))
    add_box(s, Inches(0.7), Inches(1.15), Inches(3.6), Inches(0.32), SOFT)
    add_text(s, Inches(0.78), Inches(1.16), Inches(3.45), Inches(0.3), "AI Challenge для ОЦО и бэкофиса", 11, ACCENT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.7), Inches(1.65), Inches(7.6), Inches(1.6), "Посмотрите на привычный процесс по-новому", 28, INK, font=HEAD)
    add_text(
        s,
        Inches(0.7),
        Inches(3.4),
        Inches(7.4),
        Inches(1.3),
        "Выберите знакомый рабочий процесс и предложите, как сделать его быстрее, удобнее или экономичнее. Создайте работающий прототип и покажите, что изменилось.",
        16,
        INK,
    )
    add_text(
        s,
        Inches(0.7),
        Inches(4.8),
        Inches(7.4),
        Inches(0.9),
        "Можно переосмыслить весь процесс или улучшить отдельный этап — например, подключить AI-агента.",
        14,
        MUTED,
    )
    add_text(s, Inches(0.7), Inches(6.85), Inches(5), Inches(0.3), "8 октября — 11 ноября", 12, MUTED)


def route(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Маршрут", "Как проходит AI Challenge", logo_grey(), "02")
    add_box(s, Inches(0.55), Inches(1.2), Inches(12.2), Inches(0.55), SOFT)
    add_text(
        s,
        Inches(0.7),
        Inches(1.28),
        Inches(11.9),
        Inches(0.4),
        "К финалу — работающий прототип: покажите, какую задачу он решает, как работает и что улучшает.",
        13,
        INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    steps = [
        ("01", "9 октября", "Обучение", "Знакомимся с подходом и инструментами"),
        ("02", "12–16 октября", "Практика", "Собираем свой минипроцесс в Битриксе на тестовом стенде, пока без агентов"),
        ("03", "19–26 октября", "Идея", "Кто прошёл в общий тур, готовит идею для челленджа"),
        ("04", "26 октября", "Представление идей и отбор", "Показываем замысел, жюри выбирает проекты в MVP"),
        ("05", "27 октября — 10 ноября", "Создание MVP", "Воплощаем идею в работающем прототипе"),
        ("06", "11 ноября", "Демодень", "Показываем решения и выбираем победителей"),
    ]
    gap, cw, ch = Inches(0.18), Inches(3.85), Inches(2.15)
    x0, y0 = Inches(0.55), Inches(1.95)
    for i, (num, date, name, sub) in enumerate(steps):
        col, row = i % 3, i // 3
        x = x0 + col * (cw + gap + Inches(0.12))
        y = y0 + row * (ch + Inches(0.22))
        card(s, x, y, cw, ch)
        add_text(s, x + Inches(0.16), y + Inches(0.16), Inches(0.55), Inches(0.4), num, 18, ACCENT, font=HEAD)
        add_text(s, x + Inches(0.72), y + Inches(0.24), Inches(2.9), Inches(0.3), date, 12, MUTED, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.16), y + Inches(0.62), cw - Inches(0.32), Inches(0.55), name, 18, INK, font=HEAD)
        add_text(s, x + Inches(0.16), y + Inches(1.2), cw - Inches(0.32), Inches(0.8), sub, 12, MUTED)


def stages(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Этапы", "От знакомства с инструментами до работающего решения", logo_grey(), "03")
    items = [
        ("01", "Обучение", "9 октября", "Разбираем, как находить возможности для улучшения процессов и какие инструменты помогут реализовать идею."),
        ("02", "Практика", "12–16 октября", "Каждый собирает свой минипроцесс в Битриксе на тестовом стенде. На этом шаге работаем без агентов."),
        ("03", "Идея", "19–26 октября", "Выбираете процесс и предлагаете, как выстроить его иначе. Можно свой, кросс-функциональный или чужой — если понимаете, как он устроен."),
        ("04", "Представление идей", "26 октября", "Коротко показываем заявку. Жюри — топ-менеджмент. Дальше идут те, кто набрал наибольшее количество баллов."),
        ("05", "MVP", "27 октября — 10 ноября", "Реализуем основной сценарий: прототип выполняет ключевую задачу, и его работу можно показать. Полная готовность к внедрению не требуется."),
        ("06", "Демодень", "11 ноября", "Показываем прототип, сравниваем «до» и «после», рассказываем о результате. Выбираем победителей."),
    ]
    gap, cw, ch = Inches(0.18), Inches(3.85), Inches(2.45)
    x0, y0 = Inches(0.55), Inches(1.25)
    for i, (num, name, date, text) in enumerate(items):
        col, row = i % 3, i // 3
        x = x0 + col * (cw + gap + Inches(0.12))
        y = y0 + row * (ch + Inches(0.18))
        card(s, x, y, cw, ch)
        add_text(s, x + Inches(0.16), y + Inches(0.14), Inches(1), Inches(0.28), num, 11, ACCENT)
        add_text(s, x + Inches(0.16), y + Inches(0.4), Inches(2.3), Inches(0.4), name, 16, INK, font=HEAD)
        add_text(s, x + Inches(2.35), y + Inches(0.46), Inches(1.35), Inches(0.3), date, 11, MUTED, align=PP_ALIGN.RIGHT)
        add_text(s, x + Inches(0.16), y + Inches(0.9), cw - Inches(0.32), Inches(1.4), text, 12, INK)


def idea(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Этап идея · 19–26 октября", "Что писать в заявке и как оцениваем идеи", logo_grey(), "04")
    add_text(
        s,
        Inches(0.55),
        Inches(1.15),
        Inches(12.2),
        Inches(0.7),
        "Выбираете процесс и предлагаете, как выстроить его иначе. Можно свой, кросс-функциональный или процесс другой функции — если понимаете, как он устроен. Можно отказаться от процесса или его части, если ясно, как тогда решается задача.",
        13,
        INK,
    )
    card(s, Inches(0.55), Inches(1.9), Inches(7.15), Inches(2.85))
    add_text(s, Inches(0.7), Inches(2.0), Inches(3), Inches(0.28), "В заявке коротко", 11, ACCENT)
    points = [
        "1. Какой процесс меняем и для кого он существует.",
        "2. Как он работает сейчас.",
        "3. Что в нём хотим изменить.",
        "4. Как предлагается выстроить новый процесс.",
        "5. За счёт чего это можно сделать: Bitrix, агенты, автоматизация. Детальной архитектуры не нужно.",
        "6. Какой эффект ожидаем: время, деньги, качество, ошибки, удобство, масштаб.",
        "7. Как будем измерять результат: было X, должно стать Y.",
        "8. Кто автор.",
    ]
    left, right = points[:4], points[4:]
    for i, t in enumerate(left):
        add_text(s, Inches(0.7), Inches(2.32) + Inches(0.34) * i, Inches(3.35), Inches(0.34), t, 11, INK)
    for i, t in enumerate(right):
        add_text(s, Inches(4.15), Inches(2.32) + Inches(0.5) * i, Inches(3.4), Inches(0.5), t, 11, INK)

    add_box(s, Inches(7.9), Inches(1.9), Inches(4.85), Inches(0.75), SOFT)
    add_text(s, Inches(8.05), Inches(1.98), Inches(4.55), Inches(0.6), "До заявки можно прийти к Никишевой Елене: помочь подумать и направить, не придумать решение за вас.", 12, INK)
    card(s, Inches(7.9), Inches(2.75), Inches(4.85), Inches(0.95))
    add_text(s, Inches(8.05), Inches(2.82), Inches(2), Inches(0.24), "Отбор", 11, ACCENT)
    add_text(s, Inches(8.05), Inches(3.08), Inches(4.55), Inches(0.55), "Пять критериев от 0 до 5. Дальше идут те, кто набрал наибольшее количество баллов.", 12, INK)
    card(s, Inches(7.9), Inches(3.8), Inches(4.85), Inches(0.95))
    add_text(s, Inches(8.05), Inches(3.88), Inches(2), Inches(0.24), "Жюри", 11, ACCENT)
    add_text(s, Inches(8.05), Inches(4.14), Inches(4.55), Inches(0.5), "Топ-менеджмент компании. На демодне веса другие.", 12, INK)

    kpis = [
        ("0–5", "Бизнес-эффект", "Время, деньги, результат для компании"),
        ("0–5", "Ценность", "Удобство для пользователей процесса"),
        ("0–5", "Качество", "Меньше ошибок и рисков"),
        ("0–5", "Масштаб", "Можно применять шире, чем в одном участке"),
        ("0–5", "Глубина и реализм", "Насколько сильно меняется процесс и насколько идея реалистична"),
    ]
    cw = Inches(2.36)
    x0 = Inches(0.55)
    for i, (lab, name, sub) in enumerate(kpis):
        x = x0 + i * (cw + Inches(0.1))
        card(s, x, Inches(4.95), cw, Inches(1.85))
        add_text(s, x + Inches(0.12), Inches(5.05), cw - Inches(0.2), Inches(0.28), lab, 11, ACCENT)
        add_text(s, x + Inches(0.12), Inches(5.32), cw - Inches(0.2), Inches(0.7), name, 14, INK, font=HEAD)
        add_text(s, x + Inches(0.12), Inches(6.05), cw - Inches(0.2), Inches(0.6), sub, 11, MUTED)
    add_text(s, Inches(0.55), Inches(6.88), Inches(12), Inches(0.22), "Кросс-функциональное решение получает ещё +2 балла сверх этих пяти оценок.", 11, MUTED)


def demo(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Demo Day · 11 ноября", "Как показываем решение и как выбираем победителей", logo_grey(), "05")
    add_text(s, Inches(0.55), Inches(1.15), Inches(12.2), Inches(0.4), "Формат максимально простой. Не презентация о будущем — показ того, что уже работает.", 14, INK)
    add_text(s, Inches(0.55), Inches(1.55), Inches(4), Inches(0.25), "Как показываем", 11, MUTED)
    top = [
        ("01", "Было", "Как процесс работал раньше."),
        ("02", "Стало", "Что изменили в логике."),
        ("03", "Как работает", "Живой сценарий, который можно пройти."),
        ("04", "Эффект", "Что стало лучше и как это видно."),
        ("05", "Путь в PROD", "Что нужно, чтобы вывести в работу."),
    ]
    cw = Inches(2.36)
    for i, (num, name, sub) in enumerate(top):
        x = Inches(0.55) + i * (cw + Inches(0.1))
        card(s, x, Inches(1.85), cw, Inches(2.05))
        add_text(s, x + Inches(0.12), Inches(1.95), Inches(0.5), Inches(0.25), num, 11, ACCENT)
        add_text(s, x + Inches(0.12), Inches(2.25), cw - Inches(0.24), Inches(0.7), name, 18, INK, font=HEAD)
        add_text(s, x + Inches(0.12), Inches(3.0), cw - Inches(0.24), Inches(0.75), sub, 12, INK)
    add_text(s, Inches(0.55), Inches(4.05), Inches(4), Inches(0.25), "Как оцениваем", 11, MUTED)
    bot = [
        ("30%", "Эффект", "Реальный или честно оценённый результат"),
        ("25%", "MVP работает", "Насколько реально работает прототип"),
        ("20%", "Новая логика", "Насколько хорошо перепридуман процесс"),
        ("15%", "Путь в PROD", "Можно ли довести до рабочей среды"),
        ("10%", "Демонстрация", "Насколько ясно показали решение"),
    ]
    for i, (lab, name, sub) in enumerate(bot):
        x = Inches(0.55) + i * (cw + Inches(0.1))
        card(s, x, Inches(4.35), cw, Inches(2.45), fill=SOFT)
        add_text(s, x + Inches(0.12), Inches(4.48), cw - Inches(0.24), Inches(0.55), lab, 24, ACCENT, font=HEAD)
        add_text(s, x + Inches(0.12), Inches(5.1), cw - Inches(0.24), Inches(0.55), name, 14, INK)
        add_text(s, x + Inches(0.12), Inches(5.7), cw - Inches(0.24), Inches(0.9), sub, 12, MUTED)


def prizes(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Demo Day · 11 ноября", "Призы", logo_grey(), "06")
    add_text(
        s,
        Inches(0.55),
        Inches(1.25),
        Inches(12.2),
        Inches(0.55),
        "Три места за результат на демодне. Приз не означает внедрение: решение по каждому MVP спонсоры принимают отдельно.",
        15,
        INK,
    )
    items = [
        ("1 место", "100 000 ₽", "Победитель Demo Day"),
        ("2 место", "50 000 ₽", "Второе место"),
        ("3 место", "30 000 ₽", "Третье место"),
    ]
    cw = Inches(3.9)
    for i, (lab, val, sub) in enumerate(items):
        x = Inches(0.55) + i * (cw + Inches(0.2))
        card(s, x, Inches(2.15), cw, Inches(3.4))
        add_text(s, x + Inches(0.3), Inches(2.45), cw - Inches(0.6), Inches(0.4), lab, 14, ACCENT)
        add_text(s, x + Inches(0.3), Inches(3.05), cw - Inches(0.6), Inches(1.2), val, 32, INK, font=HEAD)
        add_text(s, x + Inches(0.3), Inches(4.4), cw - Inches(0.6), Inches(0.5), sub, 14, MUTED)


def calendar(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "Сроки", "Календарь", logo_grey(), "07")
    add_text(s, Inches(0.55), Inches(1.2), Inches(12.2), Inches(0.35), "Официальное объявление — 8 октября. Демодень — 11 ноября.", 14, INK)
    rows = [
        ("Объявление", "8 октября", "Старт AI Challenge"),
        ("Обучение", "9 октября", "Знакомство с подходом и инструментами"),
        ("Тестовое задание", "12–16 октября", "Свой минипроцесс в Битриксе на тестовом стенде, без агентов"),
        ("Общий тур", "19 октября", "Объявление тех, кто идёт дальше"),
        ("Подготовка идей", "19–26 октября", "Описание процесса, идеи и ожидаемого эффекта"),
        ("Представление идей", "26 октября", "Показ замысла и работа жюри"),
        ("Создание MVP", "27 октября — 10 ноября", "Работающий прототип"),
        ("Демодень", "11 ноября", "Показ решений, призы и выбор победителей"),
    ]
    table = s.shapes.add_table(9, 3, Inches(0.55), Inches(1.65), Inches(12.2), Inches(5.2)).table
    table.columns[0].width = Inches(3.3)
    table.columns[1].width = Inches(3.3)
    table.columns[2].width = Inches(5.6)
    headers = ("Этап", "Когда", "Результат")
    for j, htxt in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = ""
        cell.fill.solid()
        cell.fill.fore_color.rgb = ACCENT
        p = cell.text_frame.paragraphs[0]
        run = p.add_run()
        set_run(run, htxt, 12, WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, row in enumerate(rows, start=1):
        bg = SOFT if i == 8 else WHITE
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            set_run(run, val, 13, INK)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def faq(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    header(s, "FAQ", "Вопросы и правила", logo_grey(), "08")
    items = [
        (0.55, 1.2, 8.05, 1.85, "Какие инструменты можно использовать?", "Cursor, AI-агентов, Битрикс24, интеграции и другие инструменты автоматизации — в рамках корпоративных правил доступа и работы с данными. AI может помогать создавать решение или быть частью процесса. Отдельный AI-агент в прототипе не обязателен."),
        (8.75, 1.2, 3.95, 1.85, "Можно участвовать командой?", "Нет. Участие только одиночное: один человек — один проект."),
        (0.55, 3.2, 3.95, 1.7, "Нужно уметь программировать?", "Нет. Важны знание процесса и готовность разбираться в инструментах."),
        (4.65, 3.2, 3.95, 1.7, "Обязательно запускать решение в работу?", "Нет. Для финала нужен работающий прототип, который можно продемонстрировать."),
        (8.75, 3.2, 3.95, 1.7, "Можно взять процесс другой команды?", "Да, если хорошо понимаете, как он устроен, и можете предложить конкретное изменение."),
        (0.55, 5.05, 12.15, 1.7, "Что будет после демодня?", "По каждому MVP отдельно решим, что делать дальше: дорабатывать, внедрять или завершить эксперимент."),
    ]
    for x, y, w, h, q, a in items:
        card(s, Inches(x), Inches(y), Inches(w), Inches(h))
        add_text(s, Inches(x + 0.16), Inches(y + 0.12), Inches(w - 0.32), Inches(0.45), q, 14, INK, font=HEAD)
        add_text(s, Inches(x + 0.16), Inches(y + 0.6), Inches(w - 0.32), Inches(h - 0.75), a, 12, MUTED)


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    cover(prs)
    route(prs)
    stages(prs)
    idea(prs)
    demo(prs)
    prizes(prs)
    calendar(prs)
    faq(prs)
    out = "AI Challenge.pptx"
    prs.save(out)
    print(out)


if __name__ == "__main__":
    main()
