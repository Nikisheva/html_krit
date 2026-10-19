#!/usr/bin/env python3
"""Short Cursor briefing for AI Challenge participants (back office and SSC)."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
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


def add_text(slide, l, t, w, h, text, size, color, font=BODY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    set_run(run, text, size, color, font=font, bold=bold)
    return sh


def card(slide, l, t, w, h, fill=WHITE):
    add_box(slide, l, t, w, h, fill, LINE)
    add_bar(slide, l, t, w)
    return l, t, w, h


def header(slide, kicker, title, num):
    add_box(slide, Inches(0.28), Inches(1.18), Pt(2.5), Inches(5.55), ACCENT)
    add_text(slide, Inches(0.55), Inches(0.22), Inches(10.4), Inches(0.32), kicker, 12, ACCENT)
    add_text(slide, Inches(0.55), Inches(0.48), Inches(10.6), Inches(0.55), title, 26, INK, font=HEAD)
    slide.shapes.add_picture("assets/pptx/logo-grey.png", Inches(11.85), Inches(0.28), height=Inches(0.42))
    add_text(slide, Inches(0.55), Inches(7.12), Inches(9), Inches(0.28), "AI Challenge · ОЦО и бэкофис", 11, ACCENT)
    add_text(slide, Inches(12.15), Inches(7.12), Inches(0.7), Inches(0.28), num, 11, MUTED, align=PP_ALIGN.RIGHT)


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_box(s, 0, 0, W, H, WHITE)
    return s


def cover(prs):
    s = blank(prs)
    builder = s.shapes.build_freeform(Inches(10.55), Inches(0))
    builder.add_line_segments([(W, Inches(0)), (W, H), (Inches(8.82), H)], close=True)
    geo = builder.convert_to_shape()
    geo.shadow.inherit = False
    geo.fill.solid()
    geo.fill.fore_color.rgb = ACCENT
    geo.line.fill.background()
    s.shapes.add_picture("assets/pptx/logo-white.png", Inches(11.85), Inches(0.32), height=Inches(0.42))
    add_box(s, Inches(0.7), Inches(1.35), Inches(4.7), Inches(0.36), SOFT)
    add_text(s, Inches(0.82), Inches(1.37), Inches(4.45), Inches(0.32), "AI Challenge · ОЦО и бэкофис", 13, ACCENT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.7), Inches(1.95), Inches(8.2), Inches(1.5), "Как работать\nв Cursor", 40, INK, font=HEAD)
    add_text(
        s,
        Inches(0.7),
        Inches(4.15),
        Inches(7.6),
        Inches(1.15),
        "Ставите задачу обычным языком. Агент готовит черновик в файлах проекта. Вы проверяете, что процесс и правда стал лучше.",
        18,
        INK,
    )
    add_text(s, Inches(0.7), Inches(6.85), Inches(6), Inches(0.3), "Обучение · 20 октября", 14, MUTED)


def what(prs):
    s = blank(prs)
    header(s, "С чего начать", "Cursor работает в папке проекта", "02")
    add_text(
        s,
        Inches(0.55),
        Inches(1.22),
        Inches(12.2),
        Inches(0.62),
        "Это не отдельный чат «просто поговорить». Откройте папку, которую выдали на обучение: агент читает файлы и может их изменить. Что оставить — решаете вы.",
        16,
        INK,
    )
    items = [
        ("01", "Говорите о процессе", "Что сейчас болит, кто делает шаг и какой результат можно показать коллегам."),
        ("02", "Агент делает черновик", "Ищет файлы, предлагает правки и вносит их. Новая тема — новый чат."),
        ("03", "Вы принимаете работу", "Откройте результат и пройдите сценарий. Не сошлось — напишите, что поправить."),
    ]
    gap, cw, ch = Inches(0.2), Inches(3.95), Inches(3.2)
    x0, y0 = Inches(0.55), Inches(2.0)
    for i, (num, title, body) in enumerate(items):
        x = x0 + i * (cw + gap)
        card(s, x, y0, cw, ch)
        add_text(s, x + Inches(0.28), y0 + Inches(0.35), Inches(3.3), Inches(0.5), num, 22, ACCENT, font=HEAD)
        add_text(s, x + Inches(0.28), y0 + Inches(1.05), Inches(3.4), Inches(0.8), title, 20, INK, font=HEAD)
        add_text(s, x + Inches(0.28), y0 + Inches(1.85), Inches(3.4), Inches(1.1), body, 15, MUTED)
    add_box(s, Inches(0.55), Inches(5.42), Inches(12.22), Inches(1.1), SOFT)
    add_text(
        s,
        Inches(0.78),
        Inches(5.55),
        Inches(11.8),
        Inches(0.85),
        "На тестовом туре агенты не используются: там свой минипроцесс в Битриксе. Cursor берите, чтобы разобрать задачу и собрать MVP.",
        16,
        INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def task(prs):
    s = blank(prs)
    header(s, "Главный навык", "Одна задача — четыре части", "03")
    parts = [
        ("01", "Как сейчас", "Процесс своими словами: кто, какой шаг, где теряется время."),
        ("02", "Что получить", "Что должно появиться: форма, статус, маршрут, письмо, отчёт."),
        ("03", "Чего нельзя", "Боевые данные, лишние системы, то, что трогать не просили."),
        ("04", "Когда готово", "Один сценарий, который вы сами можете пройти от начала до конца."),
    ]
    cw = Inches(2.95)
    gap = Inches(0.16)
    x0, y0, ch = Inches(0.55), Inches(1.28), Inches(2.15)
    for i, (num, title, body) in enumerate(parts):
        x = x0 + i * (cw + gap)
        card(s, x, y0, cw, ch, SOFT)
        add_text(s, x + Inches(0.16), y0 + Inches(0.18), Inches(2.6), Inches(0.32), num, 14, ACCENT, font=HEAD)
        add_text(s, x + Inches(0.16), y0 + Inches(0.5), Inches(2.6), Inches(0.4), title, 16, INK, font=HEAD)
        add_text(s, x + Inches(0.16), y0 + Inches(1.0), Inches(2.62), Inches(0.95), body, 13, MUTED)

    card(s, Inches(0.55), Inches(3.62), Inches(12.22), Inches(2.05), WHITE)
    add_text(s, Inches(0.78), Inches(3.78), Inches(2), Inches(0.3), "Так", 13, ACCENT, bold=True)
    add_text(
        s,
        Inches(0.78),
        Inches(4.12),
        Inches(11.75),
        Inches(1.35),
        "«Согласование счёта идёт письмами и таблицей. Собери прототип: заявка, маршрут из трёх шагов и статус. Боевые данные не используй. Готово, когда я пройду путь от заявки до „согласовано“.»",
        16,
        INK,
    )
    add_text(
        s,
        Inches(0.55),
        Inches(5.85),
        Inches(12.2),
        Inches(0.95),
        "Не так: «Сделай автоматизацию». Новая тема — новый чат. Если вышло не то, напишите, что именно поправить. Нужный файл укажите через @.",
        14,
        MUTED,
    )


def modes(prs):
    s = blank(prs)
    header(s, "Панель агента", "Три режима, больше пока не нужно", "04")
    add_text(
        s,
        Inches(0.55),
        Inches(1.22),
        Inches(12.2),
        Inches(0.55),
        "Открыть панель: Ctrl+I в Windows, Cmd+I на Mac. Режим — внизу панели или Shift+Tab.",
        16,
        INK,
    )
    items = [
        ("Agent", "Собирает и правит", "Основной режим. Просите сделать форму, маршрут, текст, правку. Смотрите, какие файлы он затронул."),
        ("Ask", "Только объясняет", "Файлы не меняет. Спрашивайте, как устроен шаг, что означает поле, где искать. Удобно, прежде чем что-то переписывать."),
        ("Plan", "Сначала план", "Для большой задачи. Агент уточняет, пишет план, ждёт вашего «да» и только потом собирает. План можно поправить."),
    ]
    cw, ch = Inches(3.95), Inches(3.2)
    gap = Inches(0.2)
    y0 = Inches(1.95)
    for i, (name, title, body) in enumerate(items):
        x = Inches(0.55) + i * (cw + gap)
        card(s, x, y0, cw, ch)
        add_text(s, x + Inches(0.28), y0 + Inches(0.32), Inches(3.4), Inches(0.4), name, 14, ACCENT, bold=True)
        add_text(s, x + Inches(0.28), y0 + Inches(0.85), Inches(3.4), Inches(0.7), title, 22, INK, font=HEAD)
        add_text(s, x + Inches(0.28), y0 + Inches(1.7), Inches(3.4), Inches(1.3), body, 15, MUTED)
    add_text(
        s,
        Inches(0.55),
        Inches(5.4),
        Inches(12.2),
        Inches(1.2),
        "Не знаете, с чего начать: откройте Ask и попросите объяснить задачу обычным языком. Когда понятно, что должно получиться, переключайтесь в Plan или сразу в Agent.",
        16,
        INK,
    )


def rules(prs):
    s = blank(prs)
    header(s, "Границы", "Четыре правила на всё время челленджа", "05")
    items = [
        ("01", "Секреты", "Пароли, ключи доступа и коды из писем в сообщение не вставляем. Даже «на один раз» и даже если агент просит."),
        ("02", "Боевые данные", "Персональные и финансовые данные, выгрузки из рабочих систем и реальные договоры в проект и в чат не выносим."),
        ("03", "Только Dev", "Прототип собираем на тестовых стендах: Bitrix Dev и Work IT Dev. На рабочий контур сами не переносим."),
        ("04", "Проверка", "Результат открываете сами и проходите сценарий. Доступы выдаёт ИТ, права в обход не настраиваем."),
    ]
    positions = [(0, 0), (1, 0), (0, 1), (1, 1)]
    cw, ch = Inches(6.0), Inches(2.35)
    gap_x, gap_y = Inches(0.22), Inches(0.18)
    x0, y0 = Inches(0.55), Inches(1.28)
    for (num, title, body), (col, row) in zip(items, positions):
        x = x0 + col * (cw + gap_x)
        y = y0 + row * (ch + gap_y)
        card(s, x, y, cw, ch)
        add_text(s, x + Inches(0.28), y + Inches(0.28), Inches(0.7), Inches(0.4), num, 18, ACCENT, font=HEAD)
        add_text(s, x + Inches(1.05), y + Inches(0.3), Inches(4.6), Inches(0.4), title, 20, INK, font=HEAD)
        add_text(s, x + Inches(0.28), y + Inches(0.95), Inches(5.45), Inches(1.15), body, 15, MUTED)


def quote(prs):
    s = blank(prs)
    builder = s.shapes.build_freeform(Inches(10.55), Inches(0))
    builder.add_line_segments([(W, Inches(0)), (W, H), (Inches(8.82), H)], close=True)
    geo = builder.convert_to_shape()
    geo.shadow.inherit = False
    geo.fill.solid()
    geo.fill.fore_color.rgb = ACCENT
    geo.line.fill.background()
    s.shapes.add_picture("assets/pptx/logo-white.png", Inches(11.85), Inches(0.32), height=Inches(0.42))
    add_text(s, Inches(0.7), Inches(1.15), Inches(8), Inches(0.35), "Вместо заключения", 14, ACCENT)
    add_text(
        s,
        Inches(0.7),
        Inches(1.7),
        Inches(8.3),
        Inches(2.6),
        "«Будущее уже здесь.\nПросто оно распределено\nнеравномерно.»",
        32,
        INK,
        font=HEAD,
    )
    add_text(s, Inches(0.7), Inches(4.55), Inches(7.5), Inches(0.35), "Уильям Гибсон", 16, INK)
    add_text(s, Inches(0.7), Inches(4.95), Inches(7.8), Inches(0.3), "The Economist, 4 декабря 2003. Перевод.", 13, MUTED)
    add_text(
        s,
        Inches(0.7),
        Inches(5.7),
        Inches(7.8),
        Inches(0.9),
        "На AI Challenge это ваш знакомый процесс: будущее, которое можно открыть и показать.",
        16,
        INK,
    )


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    prs.core_properties.title = "Как работать в Cursor"
    prs.core_properties.subject = "AI Challenge · ОЦО и бэкофис"
    cover(prs)
    what(prs)
    task(prs)
    modes(prs)
    rules(prs)
    quote(prs)
    out = "Cursor для AI Challenge.pptx"
    prs.save(out)
    print(out)


if __name__ == "__main__":
    main()
