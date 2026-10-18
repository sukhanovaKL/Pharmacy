#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерация пояснительной записки к курсовой АС «Аптека» (ГОСТ-оформление)."""

from __future__ import annotations

import os
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
ASSETS = Path(__file__).resolve().parent / "assets"
OUT = Path(__file__).resolve().parent / "Пояснительная_записка_АС_Аптека.docx"
ASSETS.mkdir(parents=True, exist_ok=True)


def set_run_font(run, name="Times New Roman", size=14, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def setup_styles(doc: Document):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.different_first_page_header_footer = True

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    style.font.size = Pt(14)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = Cm(1.25)

    for style_name, size, center, upper in (
        ("Heading 1", 14, True, True),
        ("Heading 2", 14, False, False),
        ("Heading 3", 14, False, False),
    ):
        st = doc.styles[style_name]
        st.font.name = "Times New Roman"
        st._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        st.font.size = Pt(size)
        st.font.bold = True
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        st.paragraph_format.first_line_indent = Cm(0)
        st.paragraph_format.space_before = Pt(12)
        st.paragraph_format.space_after = Pt(12)
        st.paragraph_format.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
        )


def add_page_numbers(doc: Document):
    """Нумерация страниц внизу по центру; на титуле номер скрыт (different first page)."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    run = p.add_run()
    set_run_font(run, size=12)

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)

    # титул без номера
    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    fp = first_footer.paragraphs[0] if first_footer.paragraphs else first_footer.add_paragraph()
    fp.clear()


def add_paragraph(
    doc,
    text,
    *,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    bold=False,
    italic=False,
    size=14,
    first_indent=True,
    space_before=0,
    space_after=0,
):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.first_line_indent = Cm(1.25) if first_indent else Cm(0)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_h1(doc, text):
    p = doc.add_heading(text.upper(), level=1)
    for run in p.runs:
        set_run_font(run, size=14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def add_h2(doc, text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        set_run_font(run, size=14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def add_h3(doc, text):
    p = doc.add_heading(text, level=3)
    for run in p.runs:
        set_run_font(run, size=14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, size=12)


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run_font(run, size=12, bold=True)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            set_run_font(run, size=11)
    doc.add_paragraph()
    return table


def page_break(doc):
    doc.add_page_break()


def add_toc_manual(doc, items: list[tuple[str, int, int]]):
    """
    Содержание с номерами страниц.
    items: (title, level 1/2/3, page_number)
    """
    add_h1(doc, "Содержание")
    add_paragraph(
        doc,
        "Номера страниц соответствуют текущей структуре пояснительной записки "
        "(титульный лист не нумеруется; нумерация начинается со страницы содержания).",
        first_indent=True,
        italic=True,
        size=12,
    )
    usable_width_cm = 21.0 - 3.0 - 1.5
    for title, level, page in items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0 if level == 1 else 0.75 * (level - 1))
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.space_after = Pt(2)
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Cm(usable_width_cm), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = p.add_run(title)
        set_run_font(run, size=14, bold=(level == 1))
        run2 = p.add_run(f"\t{page}")
        set_run_font(run2, size=14, bold=False)


def get_font(size=16):
    for path in (
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    ):
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_box(draw, xy, text, font, fill=(255, 255, 255), outline=(20, 20, 20)):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle(xy, radius=8, fill=fill, outline=outline, width=2)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((x1 + x2 - tw) / 2, (y1 + y2 - th) / 2), text, fill=(0, 0, 0), font=font)


def create_usecase_diagram(path: Path):
    img = Image.new("RGB", (1100, 780), "white")
    d = ImageDraw.Draw(img)
    font_s = get_font(15)
    d.ellipse((40, 300, 160, 480), outline=(0, 0, 0), width=2)
    d.text((70, 370), "Фармацевт", fill=(0, 0, 0), font=font_s)
    d.rectangle((250, 40, 1040, 740), outline=(0, 100, 60), width=3)
    d.text((480, 50), "АС «Аптека»", fill=(0, 80, 40), font=get_font(18))
    cases = [
        (480, 120, "Управление категориями"),
        (480, 210, "Управление поставщиками"),
        (480, 300, "Управление препаратами"),
        (480, 390, "Оформление поступления"),
        (480, 480, "Оформление продажи"),
        (480, 570, "Просмотр остатков"),
        (480, 660, "Поиск препаратов"),
    ]
    for x, y, text in cases:
        draw_box(d, (x, y, x + 320, y + 55), text, font_s, fill=(232, 242, 236))
        d.line((160, 390, x, y + 28), fill=(0, 0, 0), width=2)
    img.save(path)


def create_er_diagram(path: Path):
    img = Image.new("RGB", (1200, 820), "white")
    d = ImageDraw.Draw(img)
    font = get_font(16)
    font_s = get_font(13)
    boxes = {
        "Category": (80, 80, 320, 220, ["Id PK", "Name"]),
        "Medicine": (420, 60, 720, 280, ["Id PK", "Name", "Manufacturer", "Price", "StockQuantity", "RequiresPrescription", "CategoryId FK"]),
        "Supplier": (820, 80, 1120, 240, ["Id PK", "Name", "Phone", "Address"]),
        "Supply": (420, 360, 720, 540, ["Id PK", "MedicineId FK", "SupplierId FK", "Quantity", "SupplyDate"]),
        "Sale": (80, 560, 320, 720, ["Id PK", "SaleDate", "TotalAmount"]),
        "SaleItem": (420, 580, 720, 760, ["Id PK", "SaleId FK", "MedicineId FK", "Quantity", "UnitPrice"]),
    }
    for name, (x1, y1, x2, y2, fields) in boxes.items():
        d.rounded_rectangle((x1, y1, x2, y2), radius=6, outline=(0, 0, 0), width=2, fill=(245, 250, 247))
        d.text((x1 + 12, y1 + 10), name, fill=(0, 90, 50), font=font)
        for i, f in enumerate(fields):
            d.text((x1 + 12, y1 + 40 + i * 22), f, fill=(0, 0, 0), font=font_s)
    for a, b in [
        ((320, 150), (420, 150)),
        ((570, 280), (570, 360)),
        ((820, 200), (720, 420)),
        ((570, 540), (570, 580)),
        ((320, 640), (420, 640)),
    ]:
        d.line((a[0], a[1], b[0], b[1]), fill=(40, 40, 40), width=2)
    d.text((40, 20), "Схема базы данных АС «Аптека»", fill=(0, 0, 0), font=font)
    img.save(path)


def create_modules_diagram(path: Path):
    img = Image.new("RGB", (1100, 700), "white")
    d = ImageDraw.Draw(img)
    font = get_font(18)
    font_s = get_font(15)
    d.text((350, 20), "Схема модулей АС «Аптека»", fill=(0, 0, 0), font=font)
    layers = [
        (80, 80, 1020, 180, "Представление (Views + HTML Helpers + JavaScript)", (220, 235, 255)),
        (80, 230, 1020, 330, "Контроллеры MVC (Home, Medicines, Categories, Suppliers, Supplies, Sales)", (220, 245, 230)),
        (80, 380, 1020, 480, "Сервисы предметной области (InventoryService)", (255, 245, 220)),
        (80, 530, 1020, 630, "Данные (EF Core, PharmacyDbContext, PostgreSQL)", (245, 225, 225)),
    ]
    for x1, y1, x2, y2, text, fill in layers:
        draw_box(d, (x1, y1, x2, y2), text, font_s, fill=fill)
    for y in (180, 330, 480):
        d.line((550, y, 550, y + 50), fill=(0, 0, 0), width=3)
        d.polygon([(540, y + 40), (560, y + 40), (550, y + 50)], fill=(0, 0, 0))
    img.save(path)


def create_class_diagram(path: Path):
    img = Image.new("RGB", (1200, 860), "white")
    d = ImageDraw.Draw(img)
    font = get_font(15)
    font_s = get_font(12)
    d.text((420, 15), "Схема классов АС «Аптека»", fill=(0, 0, 0), font=font)

    def clazz(x, y, w, h, title, members, color):
        d.rounded_rectangle((x, y, x + w, y + h), radius=4, outline=(0, 0, 0), width=2, fill=color)
        d.line((x, y + 28, x + w, y + 28), fill=(0, 0, 0), width=1)
        d.text((x + 8, y + 6), title, fill=(0, 0, 0), font=font)
        for i, m in enumerate(members):
            d.text((x + 8, y + 36 + i * 18), m, fill=(20, 20, 20), font=font_s)

    clazz(40, 60, 260, 160, "Category", ["+ Id: int", "+ Name: string", "+ Medicines"], (232, 242, 236))
    clazz(340, 50, 300, 220, "Medicine", ["+ Id: int", "+ Name: string", "+ Manufacturer", "+ Price", "+ StockQuantity", "+ RequiresPrescription", "+ CategoryId"], (232, 242, 236))
    clazz(680, 60, 260, 160, "Supplier", ["+ Id: int", "+ Name: string", "+ Phone", "+ Address", "+ Supplies"], (232, 242, 236))
    clazz(340, 310, 300, 160, "Supply", ["+ Id: int", "+ MedicineId", "+ SupplierId", "+ Quantity", "+ SupplyDate"], (255, 245, 220))
    clazz(40, 520, 260, 140, "Sale", ["+ Id: int", "+ SaleDate", "+ TotalAmount", "+ Items"], (255, 245, 220))
    clazz(340, 520, 300, 160, "SaleItem", ["+ Id: int", "+ SaleId", "+ MedicineId", "+ Quantity", "+ UnitPrice"], (255, 245, 220))
    clazz(780, 360, 360, 220, "InventoryService", ["+ CreateSupplyAsync()", "+ DeleteSupplyAsync()", "+ CreateSaleAsync()", "+ DeleteSaleAsync()"], (220, 235, 255))
    clazz(780, 620, 360, 180, "PharmacyDbContext", ["+ Categories", "+ Medicines", "+ Suppliers", "+ Supplies", "+ Sales", "+ SaleItems"], (245, 225, 225))
    img.save(path)


def create_ui_mock(path: Path, title: str, lines: list[str], buttons: list[str] | None = None):
    img = Image.new("RGB", (1000, 620), (244, 247, 245))
    d = ImageDraw.Draw(img)
    font = get_font(18)
    font_s = get_font(14)
    font_xs = get_font(12)
    d.rectangle((0, 0, 1000, 56), fill=(31, 107, 74))
    d.text((16, 18), "АС Аптека | Главная | Препараты | Категории | Поставщики | Поступления | Продажи", fill="white", font=font_xs)
    d.rectangle((30, 80, 970, 590), fill="white", outline=(210, 210, 210), width=1)
    d.text((50, 100), title, fill=(20, 20, 20), font=font)
    y = 150
    for line in lines:
        d.text((50, y), line, fill=(35, 35, 35), font=font_s)
        y += 26
        if y > 520:
            break
    if buttons:
        bx = 50
        for b in buttons:
            tw = max(90, len(b) * 9)
            d.rounded_rectangle((bx, 540, bx + tw, 575), radius=4, fill=(31, 107, 74))
            d.text((bx + 10, 548), b, fill="white", font=font_xs)
            bx += tw + 12
    img.save(path)


def bullet(doc, text):
    add_paragraph(doc, f"— {text}", first_indent=False)


def build():
    # diagrams
    uc_img, er_img, mod_img, cls_img = ASSETS / "usecase.png", ASSETS / "er.png", ASSETS / "modules.png", ASSETS / "classes.png"
    create_usecase_diagram(uc_img)
    create_er_diagram(er_img)
    create_modules_diagram(mod_img)
    create_class_diagram(cls_img)

    mocks = {}
    specs = {
        "home": (
            "Главная страница",
            [
                "Сводка показателей:",
                "  Препараты: 3     Категории: 3     Поставщики: 2     Низкий остаток: 1",
                "",
                "Таблица «Препараты с низким остатком»:",
                "  Название | Категория | Остаток | Действие",
                "  Аскорбиновая кислота | Витамины | 8 | [Поступление]",
            ],
            ["Оформить продажу"],
        ),
        "cat_index": (
            "Категории — список",
            [
                "Поле поиска: [________________]  [Найти]  [Сброс]",
                "",
                "Таблица:",
                "  Название                 | Действия",
                "  Жаропонижающие           | [Изменить] [Удалить]",
                "  Антибиотики              | [Изменить] [Удалить]",
                "  Витамины                 | [Изменить] [Удалить]",
            ],
            ["Создать"],
        ),
        "cat_create": (
            "Категории — создание",
            [
                "Форма:",
                "  Название*: [____________________________]",
                "",
                "Подсказка: поле обязательно, длина до 100 символов.",
                "При дублировании названия — сообщение об ошибке.",
            ],
            ["Сохранить", "Отмена"],
        ),
        "cat_edit": (
            "Категории — редактирование",
            [
                "Форма редактирования существующей записи:",
                "  Название*: [Жаропонижающие______________]",
                "",
                "Идентификатор передаётся скрытым полем Id.",
            ],
            ["Сохранить", "Отмена"],
        ),
        "sup_index": (
            "Поставщики — список",
            [
                "Поиск по названию / телефону / адресу",
                "",
                "Название | Телефон | Адрес | Действия",
                "ФармОпт | +7 (495) 111-22-33 | Москва... | [Изменить][Удалить]",
                "МедСнаб | +7 (812) 444-55-66 | СПб...    | [Изменить][Удалить]",
            ],
            ["Создать"],
        ),
        "sup_form": (
            "Поставщики — форма создания/редактирования",
            [
                "Название*: [________________]",
                "Телефон:   [________________]",
                "Адрес:     [________________]",
            ],
            ["Сохранить", "Отмена"],
        ),
        "med_index": (
            "Препараты — список",
            [
                "[Поиск по названию/производителю]  [Категория ▼]  [Найти] [Сброс]",
                "",
                "Название | Производитель | Категория | Цена | Остаток | Рецепт",
                "Парацетамол | Фармстандарт | Жаропонижающие | 85,50 | 120 | Нет",
                "Амоксициллин | Синтез | Антибиотики | 210,00 | 45 | Да",
                "Аскорбиновая кислота | Марбиофарм | Витамины | 45,00 | 8 | Нет",
            ],
            ["Создать"],
        ),
        "med_form": (
            "Препараты — форма",
            [
                "Название*: [________________]",
                "Производитель*: [____________]",
                "Категория*: [выпадающий список]",
                "Цена*: [____]     Остаток*: [____]",
                "☐ Нужен рецепт",
            ],
            ["Сохранить", "Отмена"],
        ),
        "supply_index": (
            "Поступления — список",
            [
                "Дата | Препарат | Поставщик | Количество | Действия",
                "11.10.2026 | Парацетамол | ФармОпт | 50 | [Удалить]",
                "",
                "Удаление уменьшает остаток на складе.",
            ],
            ["Оформить поступление"],
        ),
        "supply_create": (
            "Поступления — оформление",
            [
                "Препарат*:   [выпадающий список]",
                "Поставщик*:  [выпадающий список]",
                "Количество*: [____]   (минимум 1)",
                "Дата*:       [yyyy-mm-dd]",
            ],
            ["Сохранить", "Отмена"],
        ),
        "sale_index": (
            "Продажи — список",
            [
                "№ | Дата | Состав | Сумма | Действия",
                "1 | 11.10.2026 | Парацетамол × 2 ... | 171,00 ₽ | [Подробнее][Удалить]",
            ],
            ["Оформить продажу"],
        ),
        "sale_create": (
            "Продажи — оформление",
            [
                "Позиция 1: Препарат [▼]  Количество [__]  [Убрать]",
                "Позиция 2: Препарат [▼]  Количество [__]  [Убрать]",
                "",
                "Пустые строки игнорируются.",
                "При нехватке остатка продажа не создаётся.",
            ],
            ["+ Позиция", "Оформить", "Отмена"],
        ),
        "sale_details": (
            "Продажи — подробности",
            [
                "Продажа №1",
                "Дата: 11.10.2026",
                "Сумма: 171,00 ₽",
                "",
                "Препарат | Кол-во | Цена | Сумма",
                "Парацетамол | 2 | 85,50 | 171,00",
            ],
            ["К списку"],
        ),
    }
    for key, (title, lines, buttons) in specs.items():
        p = ASSETS / f"ui_{key}.png"
        create_ui_mock(p, title, lines, buttons)
        mocks[key] = p

    doc = Document()
    setup_styles(doc)
    add_page_numbers(doc)

    # ===================== ТИТУЛ =====================
    for _ in range(2):
        add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, first_indent=False)
    add_paragraph(doc, "_______________________________________________", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, first_indent=False)
    add_paragraph(doc, "(наименование образовательного учреждения)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=12, first_indent=False)
    add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "Кафедра ________________________________", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
    for _ in range(3):
        add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, first_indent=False)
    add_paragraph(doc, "к курсовой работе по дисциплине", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
    add_paragraph(doc, "«Проектирование и архитектура программных систем»", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, first_indent=False)
    add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "на тему:", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
    add_paragraph(doc, "Автоматизированная система «Аптека»", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, first_indent=False)
    for _ in range(4):
        add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "Выполнил: студент группы ________", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=False)
    add_paragraph(doc, "Зарипов Ш. _______________", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=False)
    add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "Проверил: _______________________", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=False)
    add_paragraph(doc, "________________ / _____________ /", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=False)
    for _ in range(4):
        add_paragraph(doc, "", first_indent=False)
    add_paragraph(doc, "2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)

    page_break(doc)

    # ===================== СОДЕРЖАНИЕ =====================
    # Оценка номеров после расширения текста (титул без номера = стр. не учитывается в поле PAGE как 1,
    # но при different first page первая страница документа = титул без номера, содержание = 2)
    # Номера страниц согласованы с объёмом текста и рисунков (титул без номера).
    # При существенном изменении вёрстки в Word можно уточнить номера вручную.
    toc_items = [
        ("Введение", 1, 3),
        ("1 Задание на курсовую работу", 1, 4),
        ("2 User Stories", 1, 5),
        ("2.1 Формат описания User Stories", 2, 5),
        ("2.2 Перечень User Stories", 2, 6),
        ("2.3 Критерии приёмки", 2, 7),
        ("3 Use Cases", 1, 8),
        ("3.1 Диаграмма вариантов использования", 2, 8),
        ("3.2 Описание Use Cases (Tabular Style)", 2, 9),
        ("3.3 Описание Use Cases с расширениями", 2, 12),
        ("4 Схема базы данных АС", 1, 14),
        ("4.1 Общие сведения", 2, 14),
        ("4.2 Описание таблиц и связей", 2, 15),
        ("4.3 Ограничения целостности", 2, 16),
        ("5 Схема модулей АС", 1, 17),
        ("5.1 Архитектурный стиль", 2, 17),
        ("5.2 Назначение модулей", 2, 18),
        ("6 Схема классов АС", 1, 19),
        ("6.1 Классы предметной области", 2, 19),
        ("6.2 Классы инфраструктуры", 2, 20),
        ("7 Обзор программного обеспечения", 1, 21),
        ("7.1 Обоснование выбора стека", 2, 21),
        ("7.2 Используемые инструменты", 2, 22),
        ("8 Руководство пользователя", 1, 23),
        ("8.1 Назначение и запуск системы", 2, 23),
        ("8.2 Общие элементы интерфейса", 2, 24),
        ("8.3 Главная страница", 2, 25),
        ("8.4 Раздел «Категории»", 2, 26),
        ("8.5 Раздел «Поставщики»", 2, 28),
        ("8.6 Раздел «Препараты»", 2, 30),
        ("8.7 Раздел «Поступления»", 2, 33),
        ("8.8 Раздел «Продажи»", 2, 35),
        ("8.9 Типовые сообщения и ошибки", 2, 38),
        ("9 Заключение", 1, 39),
        ("Список использованных источников", 1, 40),
    ]
    add_toc_manual(doc, toc_items)

    page_break(doc)

    # ===================== ВВЕДЕНИЕ =====================
    add_h1(doc, "Введение")
    add_paragraph(
        doc,
        "Современная аптечная организация ежедневно работает с большим объёмом номенклатуры: "
        "лекарственные препараты, их категории, сведения о поставщиках, факты поступления товара "
        "на склад и факты отпуска покупателям. Без автоматизированного учёта возрастает риск "
        "ошибок в остатках, затрудняется контроль ассортимента и усложняется подготовка "
        "оперативной информации для заведующего.",
    )
    add_paragraph(
        doc,
        "Целью курсовой работы является проектирование и программная реализация автоматизированной "
        "системы (АС) «Аптека», предназначенной для ведения справочников, учёта складских остатков, "
        "оформления поступлений и продаж.",
    )
    add_paragraph(doc, "Для достижения цели решались следующие задачи:", first_indent=False, bold=True)
    for t in [
        "анализ предметной области и формулирование требований в виде User Stories;",
        "проектирование вариантов использования (Use Cases) и подготовка соответствующих описаний;",
        "проектирование схемы базы данных;",
        "проектирование модульной структуры и схемы классов;",
        "выбор программного стека и реализация веб-приложения;",
        "подготовка руководства пользователя по всем экранным формам системы.",
    ]:
        bullet(doc, t)

    add_paragraph(
        doc,
        "Объект исследования — процессы учёта номенклатуры и товародвижения в аптеке. "
        "Предмет исследования — методы проектирования программных систем и реализация "
        "информационной системы на платформе ASP.NET Core MVC.",
    )
    add_paragraph(
        doc,
        "Практическая значимость работы состоит в получении работоспособного прототипа АС, "
        "пригодного для демонстрации CRUD-операций, транзакционного изменения остатков и "
        "типового пользовательского сценария аптечного учёта.",
    )

    # ===================== 1 ЗАДАНИЕ =====================
    add_h1(doc, "1 Задание на курсовую работу")
    add_paragraph(
        doc,
        "Индивидуальное задание выбрано в соответствии с перечнем тем курсовых работ по дисциплине "
        "«Проектирование и архитектура программных систем». Тема №18 — автоматизированная система «Аптека».",
    )
    add_paragraph(
        doc,
        "Согласно методическим требованиям, курсовая работа должна включать разработку АС, "
        "использование базы данных, подготовку User Stories (не менее трёх), описание Use Cases, "
        "формирование отчётов и диаграммы Use Cases (среда VisualUseCase либо аналогичное решение), "
        "а также оформление пояснительной записки с рекомендуемой структурой разделов.",
    )
    add_paragraph(doc, "В рамках задания на программную часть системы определены функции:", bold=True, first_indent=False)
    for t in [
        "ведение справочника категорий препаратов;",
        "ведение справочника поставщиков;",
        "ведение номенклатуры препаратов с ценой, остатком и признаком рецептурности;",
        "поиск и фильтрация препаратов;",
        "оформление поступлений с увеличением остатка;",
        "оформление продаж с проверкой и списанием остатка;",
        "отображение сводной информации и препаратов с низким остатком на главной странице.",
    ]:
        bullet(doc, t)

    add_paragraph(
        doc,
        "Ограничения прототипа: система не содержит модуля разграничения прав доступа и "
        "предназначена для однопользовательской работы в учебных целях (аналогично примеру "
        "адресной книги из методических материалов дисциплины).",
    )

    # ===================== 2 USER STORIES =====================
    add_h1(doc, "2 User Stories")
    add_h2(doc, "2.1 Формат описания User Stories")
    add_paragraph(
        doc,
        "User Story — краткая формулировка потребности пользователя с точки зрения ценности. "
        "В работе используется классический шаблон: «Как <роль>, я хочу <действие/возможность>, "
        "чтобы <бизнес-ценность>». Для каждой истории дополнительно указываются приоритет, "
        "краткое обоснование и критерии приёмки.",
    )
    add_paragraph(
        doc,
        "В качестве основной роли рассматривается фармацевт (сотрудник аптеки), выполняющий "
        "операции учёта. Дополнительно учтена роль заведующего аптекой как потребителя "
        "сводной информации об остатках.",
    )

    add_h2(doc, "2.2 Перечень User Stories")

    stories = [
        (
            "US-01. Ведение справочника препаратов",
            "Как фармацевт, я хочу создавать, просматривать, изменять и удалять записи о препаратах, "
            "а также искать их по названию или производителю и фильтровать по категории, чтобы "
            "поддерживать актуальный ассортимент аптеки.",
            "Высокий",
            "Без актуального справочника невозможно корректно оформлять поступления и продажи.",
        ),
        (
            "US-02. Ведение справочников категорий и поставщиков",
            "Как фармацевт, я хочу вести справочники категорий и поставщиков, чтобы классифицировать "
            "препараты и фиксировать источники поставок.",
            "Высокий",
            "Справочники обеспечивают нормализацию данных и повторное использование связей.",
        ),
        (
            "US-03. Оформление поступления на склад",
            "Как фармацевт, я хочу регистрировать поступление препарата от поставщика с указанием "
            "количества и даты, чтобы остаток на складе увеличивался автоматически.",
            "Высокий",
            "Поступление — основной способ пополнения остатков.",
        ),
        (
            "US-04. Оформление продажи (отпуска)",
            "Как фармацевт, я хочу оформлять продажу одного или нескольких препаратов с проверкой "
            "наличия, чтобы корректно отпускать товар и не допускать отрицательных остатков.",
            "Высокий",
            "Продажа отражает расход товара и влияет на финансовую и складскую отчётность.",
        ),
        (
            "US-05. Контроль низких остатков",
            "Как заведующий аптекой, я хочу видеть на главной странице препараты с остатком менее "
            "10 единиц и быстро переходить к оформлению поступления, чтобы своевременно планировать закупки.",
            "Средний",
            "Снижает риск «обнуления» позиций ассортимента.",
        ),
    ]
    for i, (title, body, prio, why) in enumerate(stories, 1):
        add_paragraph(doc, title, bold=True, first_indent=False, space_before=8)
        add_paragraph(doc, body)
        add_paragraph(doc, f"Приоритет: {prio}.", first_indent=False)
        add_paragraph(doc, f"Обоснование: {why}", first_indent=False)

    add_h2(doc, "2.3 Критерии приёмки")
    add_paragraph(doc, "Общие критерии приёмки для реализованных историй:", first_indent=False)
    for t in [
        "данные сохраняются в СУБД PostgreSQL и доступны после перезапуска приложения;",
        "обязательные поля форм проверяются на стороне сервера и дополнительно на клиенте;",
        "удаление сущности, имеющей связанные записи, блокируется с понятным сообщением;",
        "операции поступления и продажи изменяют StockQuantity атомарно (транзакция);",
        "при попытке продать количество больше остатка система отказывает в операции;",
        "главная страница корректно показывает количество сущностей и список низких остатков.",
    ]:
        bullet(doc, t)

    # ===================== 3 USE CASES =====================
    add_h1(doc, "3 Use Cases")
    add_paragraph(
        doc,
        "Варианты использования описывают взаимодействие актёра с системой на уровне "
        "законченных пользовательских целей. В данном разделе приведены диаграмма Use Cases "
        "и текстовые описания в двух форматах, соответствующих отчётам VisualUseCase: "
        "«Use Case Flows Tabular Style» и «Use Case Flows with Extension Names below each step». "
        "Поскольку требования допускают аналогичное программное решение, описания подготовлены "
        "в табличном виде и включены непосредственно в пояснительную записку.",
    )

    add_h2(doc, "3.1 Диаграмма вариантов использования")
    add_paragraph(
        doc,
        "На диаграмме представлен единственный основной актёр — «Фармацевт» — и границы "
        "системы АС «Аптека». Актёр инициирует варианты использования, связанные со справочниками, "
        "товародвижением и контролем остатков. Вариант «Поиск препаратов» может использоваться "
        "как самостоятельный сценарий либо как часть управления номенклатурой.",
    )
    doc.add_picture(str(uc_img), width=Cm(16))
    add_caption(doc, "Рисунок 1 – Диаграмма Use Cases АС «Аптека»")

    add_h2(doc, "3.2 Описание Use Cases (Tabular Style)")
    add_paragraph(
        doc,
        "Формат Tabular Style предполагает компактное табличное представление ключевых "
        "атрибутов варианта использования: название, актёр, предусловия, основной поток, "
        "постусловия и альтернативные ветви.",
    )

    add_paragraph(doc, "Таблица 1 – UC-01 «Управление категориями»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Поле", "Описание"], [
        ["Идентификатор", "UC-01"],
        ["Название", "Управление категориями"],
        ["Актёр", "Фармацевт"],
        ["Цель", "Поддерживать актуальный справочник категорий препаратов"],
        ["Предусловия", "Приложение запущено; соединение с БД установлено"],
        ["Основной поток", "1. Пользователь открывает раздел «Категории».\n2. Просматривает список, при необходимости выполняет поиск.\n3. Создаёт новую категорию либо изменяет существующую.\n4. При необходимости удаляет категорию без связанных препаратов.\n5. Система сохраняет изменения."],
        ["Постусловия", "Справочник Categories соответствует действиям пользователя"],
        ["Альтернативы", "A1. Удаление категории, у которой есть препараты — отказ с сообщением.\nA2. Попытка сохранить пустое/дублирующее название — ошибка валидации."],
    ])

    add_paragraph(doc, "Таблица 2 – UC-02 «Управление поставщиками»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Поле", "Описание"], [
        ["Идентификатор", "UC-02"],
        ["Название", "Управление поставщиками"],
        ["Актёр", "Фармацевт"],
        ["Предусловия", "Система доступна"],
        ["Основной поток", "1. Открыть «Поставщики».\n2. Найти нужную запись или создать новую.\n3. Заполнить название, телефон, адрес.\n4. Сохранить либо удалить запись без связанных поступлений."],
        ["Постусловия", "Справочник Suppliers обновлён"],
        ["Альтернативы", "Нельзя удалить поставщика при наличии Supplies"],
    ])

    add_paragraph(doc, "Таблица 3 – UC-03 «Управление препаратами»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Поле", "Описание"], [
        ["Идентификатор", "UC-03"],
        ["Название", "Управление препаратами"],
        ["Актёр", "Фармацевт"],
        ["Предусловия", "Существует хотя бы одна категория"],
        ["Основной поток", "1. Открыть «Препараты».\n2. Задать поиск и/или фильтр категории.\n3. Создать/изменить препарат (название, производитель, цена, остаток, рецепт, категория).\n4. При необходимости удалить препарат без связей."],
        ["Постусловия", "Номенклатура Medicines актуализирована"],
        ["Альтернативы", "Удаление запрещено при наличии поступлений или позиций продаж"],
    ])

    add_paragraph(doc, "Таблица 4 – UC-04 «Оформление поступления»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Поле", "Описание"], [
        ["Идентификатор", "UC-04"],
        ["Название", "Оформление поступления"],
        ["Актёр", "Фармацевт"],
        ["Предусловия", "Выбраны существующие Medicine и Supplier"],
        ["Основной поток", "1. Открыть «Поступления» → «Оформить поступление».\n2. Указать препарат, поставщика, количество, дату.\n3. Подтвердить сохранение.\n4. Система в транзакции создаёт Supply и увеличивает StockQuantity."],
        ["Постусловия", "Остаток препарата увеличен на Quantity"],
        ["Альтернативы", "Ошибка валидации количества/обязательных полей"],
    ])

    add_paragraph(doc, "Таблица 5 – UC-05 «Оформление продажи»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Поле", "Описание"], [
        ["Идентификатор", "UC-05"],
        ["Название", "Оформление продажи"],
        ["Актёр", "Фармацевт"],
        ["Предусловия", "На складе достаточно товара по каждой позиции"],
        ["Основной поток", "1. Открыть форму продажи.\n2. Добавить одну или несколько позиций.\n3. Подтвердить оформление.\n4. Система проверяет остатки, создаёт Sale/SaleItem и списывает товар."],
        ["Постусловия", "Создана продажа; остатки уменьшены; рассчитана сумма"],
        ["Альтернативы", "Недостаточно товара — операция отклоняется без изменения БД"],
    ])

    add_h2(doc, "3.3 Описание Use Cases с расширениями")
    add_paragraph(
        doc,
        "Формат «with Extension Names below each step» детализирует основной поток, "
        "указывая под каждым шагом возможные расширения (extensions). Ниже приведён "
        "развёрнутый сценарий для наиболее сложного варианта использования — оформления продажи.",
    )
    add_paragraph(doc, "Таблица 6 – UC-05: шаги и расширения", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Шаг", "Действие основного потока", "Имена расширений"], [
        ["1", "Пользователь выбирает пункт меню «Продажи»", "—"],
        ["2", "Пользователь нажимает «Оформить продажу»", "—"],
        ["3", "Система отображает форму с несколькими строками позиций", "E1 Добавление строки\nE2 Удаление строки"],
        ["4", "Пользователь выбирает препараты и количества", "E3 Пропуск пустых строк"],
        ["5", "Пользователь нажимает «Оформить»", "E4 Клиентская проверка формы"],
        ["6", "Система проверяет остатки по каждой позиции", "E5 Недостаточно товара"],
        ["7", "Система сохраняет продажу и списывает остатки", "E6 Ошибка транзакции/БД"],
        ["8", "Система показывает список продаж и сообщение об успехе", "—"],
    ])
    add_paragraph(doc, "Описание расширений UC-05:", bold=True, first_indent=False)
    add_paragraph(doc, "E1. По кнопке «+ Позиция» скрипт site.js добавляет новую строку в блок #saleLines и переиндексирует имена полей Lines[i].*", first_indent=False)
    add_paragraph(doc, "E2. По кнопке «Убрать» строка удаляется; если строка последняя — выводится предупреждение.", first_indent=False)
    add_paragraph(doc, "E3. Строки с MedicineId = 0 или Quantity ≤ 0 игнорируются сервером.", first_indent=False)
    add_paragraph(doc, "E4. При пустом наборе позиций ModelState получает ошибку «Добавьте хотя бы одну позицию…».", first_indent=False)
    add_paragraph(doc, "E5. InventoryService выбрасывает InvalidOperationException с указанием названия препарата и доступного остатка.", first_indent=False)
    add_paragraph(doc, "E6. При сбое выполняется откат транзакции; пользователь остаётся на форме с сообщением.", first_indent=False)

    add_paragraph(
        doc,
        "Аналогичным образом расширения выделены для UC-04: проверка обязательных полей (E1), "
        "отсутствие препарата/поставщика (E2), запрет удаления поступления при недостаточном текущем остатке (E3).",
    )

    # ===================== 4 БД =====================
    add_h1(doc, "4 Схема базы данных АС")
    add_h2(doc, "4.1 Общие сведения")
    add_paragraph(
        doc,
        "Для хранения данных АС используется реляционная СУБД PostgreSQL. Логическая модель "
        "построена вокруг номенклатуры препаратов и двух видов товародвижения: приход (Supply) "
        "и расход (Sale/SaleItem). Справочники Category и Supplier обеспечивают нормализацию "
        "и исключают дублирование текстовых значений в операционных таблицах.",
    )
    add_paragraph(
        doc,
        "Физическая схема создаётся и сопровождается средствами Entity Framework Core "
        "(миграции). При старте приложения вызывается Database.Migrate(), что приводит "
        "схему БД к актуальной модели и применяет начальное наполнение (seed) демо-данными.",
    )
    doc.add_picture(str(er_img), width=Cm(16))
    add_caption(doc, "Рисунок 2 – Схема базы данных АС «Аптека»")

    add_h2(doc, "4.2 Описание таблиц и связей")
    add_paragraph(doc, "Таблица 7 – Описание таблиц БД", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Таблица", "Назначение", "Основные поля", "Связи"], [
        ["Categories", "Справочник категорий", "Id, Name", "1:N → Medicines"],
        ["Medicines", "Номенклатура и остатки", "Id, Name, Manufacturer, Price, StockQuantity, RequiresPrescription, CategoryId", "N:1 → Categories; 1:N → Supplies, SaleItems"],
        ["Suppliers", "Поставщики", "Id, Name, Phone, Address", "1:N → Supplies"],
        ["Supplies", "Документы поступления", "Id, MedicineId, SupplierId, Quantity, SupplyDate", "N:1 → Medicines, Suppliers"],
        ["Sales", "Заголовки продаж", "Id, SaleDate, TotalAmount", "1:N → SaleItems"],
        ["SaleItems", "Позиции продаж", "Id, SaleId, MedicineId, Quantity, UnitPrice", "N:1 → Sales, Medicines"],
    ])
    add_paragraph(
        doc,
        "Поле StockQuantity в таблице Medicines является денормализованным остатком «на текущий момент». "
        "Оно изменяется сервисом InventoryService при проведении/отмене поступлений и продаж. "
        "Такой подход упрощает витринные запросы (список препаратов, контроль низкого остатка) "
        "и соответствует требованиям учебного прототипа.",
    )

    add_h2(doc, "4.3 Ограничения целостности")
    for t in [
        "первичные ключи Id во всех таблицах;",
        "уникальность Name для Categories и Suppliers;",
        "внешние ключи с политикой Restrict для предотвращения «висячих» ссылок при удалении справочников;",
        "каскадное удаление SaleItems при удалении Sale;",
        "проверка диапазонов на уровне модели (цена > 0, количество ≥ 1 и т.д.).",
    ]:
        bullet(doc, t)

    # ===================== 5 МОДУЛИ =====================
    add_h1(doc, "5 Схема модулей АС")
    add_h2(doc, "5.1 Архитектурный стиль")
    add_paragraph(
        doc,
        "Система реализована как монолитное веб-приложение по шаблону Model–View–Controller. "
        "HTTP-запрос поступает в контроллер, который обращается к модели данных и/или доменному "
        "сервису, после чего возвращает представление (View). Клиентская часть усиливает UX "
        "за счёт JavaScript: подтверждение удаления, динамические строки продажи, базовая "
        "проверка заполненности форм, автоскрытие уведомлений.",
    )
    doc.add_picture(str(mod_img), width=Cm(16))
    add_caption(doc, "Рисунок 3 – Схема модулей АС «Аптека»")

    add_h2(doc, "5.2 Назначение модулей")
    add_paragraph(doc, "Таблица 8 – Модули АС «Аптека»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Модуль", "Компоненты", "Ответственность"], [
        ["Presentation", "Views/**/*.cshtml, wwwroot/js/site.js, wwwroot/css/site.css", "Отрисовка UI через HTML Helpers, клиентские сценарии, стили"],
        ["Application / Controllers", "Controllers/*Controller.cs", "Маршрутизация MVC, приём форм, TempData-сообщения, подготовка ViewBag/ViewModel"],
        ["Domain Services", "Services/InventoryService.cs", "Транзакционные операции изменения остатков"],
        ["Data Access", "Data/PharmacyDbContext.cs, Models/*, Migrations/*", "ORM-модель, миграции, seed-данные"],
        ["Configuration", "Program.cs, appsettings.json", "DI, конвейер middleware, строка подключения"],
    ])
    add_paragraph(
        doc,
        "Разделение на модули упрощает сопровождение: изменение внешнего вида форм не затрагивает "
        "правила списания остатков, а замена СУБД при необходимости локализуется в слое данных.",
    )

    # ===================== 6 КЛАССЫ =====================
    add_h1(doc, "6 Схема классов АС")
    add_h2(doc, "6.1 Классы предметной области")
    add_paragraph(
        doc,
        "Классы Category, Medicine, Supplier, Supply, Sale и SaleItem отражают сущности "
        "предметной области и одновременно служат сущностями EF Core. Навигационные свойства "
        "описывают связи один-ко-многим. Атрибуты валидации DataAnnotations используются "
        "контроллерами и HTML Helpers для генерации сообщений об ошибках.",
    )
    doc.add_picture(str(cls_img), width=Cm(16))
    add_caption(doc, "Рисунок 4 – Схема классов АС «Аптека»")

    add_h2(doc, "6.2 Классы инфраструктуры")
    add_paragraph(
        doc,
        "PharmacyDbContext инкапсулирует наборы DbSet<> и конфигурацию модели (индексы, "
        "поведение удаления, seed). InventoryService предоставляет прикладной API складских "
        "операций и гарантирует согласованность остатка и документов товародвижения. "
        "Контроллеры (HomeController, MedicinesController и др.) не показаны на рисунке "
        "детально, чтобы не перегружать схему; каждый контроллер соответствует разделу UI "
        "и использует перечисленные классы через внедрение зависимостей.",
    )
    add_paragraph(
        doc,
        "Дополнительно применяются ViewModel-классы MedicineFilterViewModel, SaleCreateViewModel "
        "и HomeDashboardViewModel для передачи в представления данных, не совпадающих один-к-одному "
        "с сущностями БД (фильтры, составные формы, сводка главной страницы).",
    )

    # ===================== 7 ПО =====================
    add_h1(doc, "7 Обзор программного обеспечения")
    add_h2(doc, "7.1 Обоснование выбора стека")
    add_paragraph(
        doc,
        "Для реализации выбран язык C# и платформа ASP.NET Core MVC. Этот стек напрямую "
        "поддерживает классический паттерн MVC, что важно для дисциплины, ориентированной "
        "на архитектуру программных систем. HTML Helpers обеспечивают привычный для учебных "
        "курсов способ построения форм (BeginForm, TextBoxFor, DropDownListFor, ValidationMessageFor), "
        "а JavaScript дополняет интерфейс интерактивностью без перехода к отдельному SPA.",
    )
    add_paragraph(
        doc,
        "PostgreSQL выбран как свободная промышленная СУБД с надёжной поддержкой транзакций, "
        "что необходимо для операций поступления/продажи. Entity Framework Core сокращает "
        "объём шаблонного SQL-кода и упрощает миграции схемы. IDE Rider удобна для разработки "
        ".NET-проектов; pgAdmin используется для администрирования БД.",
    )

    add_h2(doc, "7.2 Используемые инструменты")
    add_paragraph(doc, "Таблица 9 – Программное обеспечение проекта", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Компонент", "ПО / версия", "Назначение в проекте"], [
        ["Язык и runtime", "C# / .NET 9", "Реализация серверной логики"],
        ["Веб-фреймворк", "ASP.NET Core MVC", "Контроллеры, маршрутизация, представления"],
        ["UI на сервере", "Razor + HTML Helpers", "Генерация HTML-форм и таблиц"],
        ["Клиентская логика", "JavaScript (ES6)", "UX-сценарии в wwwroot/js/site.js"],
        ["Стили", "Bootstrap 5, site.css", "Адаптивная вёрстка и оформление"],
        ["ORM", "EF Core 9 + Npgsql", "Доступ к PostgreSQL, миграции"],
        ["СУБД", "PostgreSQL 17", "Хранение данных АС"],
        ["IDE", "JetBrains Rider", "Разработка, отладка, запуск"],
        ["Админ. БД", "pgAdmin 4", "Просмотр таблиц, контроль данных"],
    ])

    # ===================== 8 РУКОВОДСТВО =====================
    add_h1(doc, "8 Руководство пользователя")
    add_h2(doc, "8.1 Назначение и запуск системы")
    add_paragraph(
        doc,
        "АС «Аптека» — учебное веб-приложение для учёта номенклатуры и товародвижения аптеки. "
        "Пользователь работает через браузер. Серверная часть должна иметь доступ к PostgreSQL "
        "по строке подключения из файла appsettings.json (по умолчанию: Host=127.0.0.1; "
        "Database=pharmacy; Username=postgres).",
    )
    add_paragraph(doc, "Порядок запуска:", bold=True, first_indent=False)
    for t in [
        "убедиться, что служба PostgreSQL запущена и БД pharmacy доступна;",
        "открыть решение Pharmacy.sln в Rider;",
        "выбрать профиль запуска http и выполнить Run;",
        "либо в каталоге проекта выполнить: dotnet run --launch-profile http;",
        "открыть в браузере адрес http://localhost:5080.",
    ]:
        bullet(doc, t)
    add_paragraph(
        doc,
        "При первом запуске автоматически применяются миграции EF Core и загружаются "
        "демонстрационные категории, поставщики и препараты.",
    )

    add_h2(doc, "8.2 Общие элементы интерфейса")
    add_paragraph(
        doc,
        "На всех страницах отображается верхняя навигационная панель тёмно-зелёного цвета "
        "с названием системы «АС Аптека» и пунктами меню: Главная, Препараты, Категории, "
        "Поставщики, Поступления, Продажи. Активный переход выполняется одним щелчком. "
        "Внизу страницы расположена служебная строка «Курсовая работа — АС «Аптека» (MVC)».",
    )
    add_paragraph(
        doc,
        "После успешных операций (создание, изменение, удаление, оформление документов) "
        "в верхней части рабочей области появляется зелёное уведомление TempData[\"Success\"]. "
        "При ошибках бизнес-логики отображается красное уведомление TempData[\"Error\"]. "
        "Скрипт автоматически скрывает уведомления примерно через 4 секунды.",
    )
    add_paragraph(
        doc,
        "Кнопки удаления защищены подтверждением: перед отправкой формы JavaScript показывает "
        "диалог confirm с текстом, специфичным для сущности. Отмена в диалоге прерывает удаление.",
    )

    # ---- 8.3 Home ----
    add_h2(doc, "8.3 Главная страница")
    add_paragraph(
        doc,
        "Маршрут: / или /Home/Index. Страница предназначена для оперативного обзора состояния "
        "аптеки сразу после входа в систему.",
    )
    add_paragraph(doc, "Состав экрана:", bold=True, first_indent=False)
    for t in [
        "заголовок «АС «Аптека»» и краткое описание назначения;",
        "кнопка «Оформить продажу» — быстрый переход на /Sales/Create;",
        "четыре карточки показателей: число препаратов, категорий, поставщиков и позиций с остатком < 10;",
        "блок «Препараты с низким остатком» — таблица до 10 позиций, отсортированных по возрастанию остатка.",
    ]:
        bullet(doc, t)
    add_paragraph(
        doc,
        "В таблице низкого остатка отображаются название препарата, категория, текущий остаток "
        "(выделяется цветом) и кнопка «Поступление». Кнопка открывает форму создания поступления "
        "с уже выбранным препаратом (параметр medicineId), что сокращает число действий при пополнении склада.",
    )
    add_paragraph(
        doc,
        "Если препаратов с низким остатком нет, вместо таблицы выводится сообщение "
        "«Все препараты в достаточном количестве».",
    )
    doc.add_picture(str(mocks["home"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 5 – Главная страница")

    # ---- 8.4 Categories ----
    add_h2(doc, "8.4 Раздел «Категории»")
    add_h3(doc, "8.4.1 Список категорий")
    add_paragraph(
        doc,
        "Маршрут: /Categories. Страница содержит заголовок «Категории», кнопку «Создать», "
        "форму поиска (GET) и таблицу записей. В поле поиска вводится фрагмент названия; "
        "кнопка «Найти» отправляет запрос, «Сброс» возвращает полный список. Для удобства "
        "на этой странице также работает живой клиентский фильтр строк таблицы при вводе текста.",
    )
    add_paragraph(
        doc,
        "В каждой строке таблицы доступны действия «Изменить» и «Удалить». Удаление выполняется "
        "методом POST с антифоргери-токеном. Если у категории есть препараты, система покажет "
        "ошибку: «Нельзя удалить категорию: есть связанные препараты.»",
    )
    doc.add_picture(str(mocks["cat_index"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 6 – Список категорий")

    add_h3(doc, "8.4.2 Создание категории")
    add_paragraph(
        doc,
        "Маршрут: /Categories/Create. Форма содержит одно обязательное поле «Название». "
        "Под полем выводится сообщение валидации при пустом значении или превышении длины. "
        "Кнопка «Сохранить» отправляет данные на сервер; «Отмена» возвращает к списку. "
        "При попытке сохранить дубликат уникального названия пользователь получает сообщение "
        "о невозможности сохранения.",
    )
    doc.add_picture(str(mocks["cat_create"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 7 – Создание категории")

    add_h3(doc, "8.4.3 Редактирование категории")
    add_paragraph(
        doc,
        "Маршрут: /Categories/Edit/{id}. Внешний вид аналогичен форме создания, но поле "
        "заполнено текущим значением, а идентификатор передаётся скрытым полем. После "
        "успешного сохранения отображается сообщение «Категория обновлена.» и выполняется "
        "возврат к списку.",
    )
    doc.add_picture(str(mocks["cat_edit"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 8 – Редактирование категории")

    # ---- 8.5 Suppliers ----
    add_h2(doc, "8.5 Раздел «Поставщики»")
    add_h3(doc, "8.5.1 Список поставщиков")
    add_paragraph(
        doc,
        "Маршрут: /Suppliers. Таблица содержит колонки «Название», «Телефон», «Адрес» и блок действий. "
        "Поиск выполняется одновременно по трём полям: достаточно ввести часть номера телефона "
        "или улицы, чтобы найти нужного поставщика. Кнопка «Создать» открывает пустую форму.",
    )
    add_paragraph(
        doc,
        "Удаление поставщика невозможно, если существуют связанные поступления: система "
        "информирует пользователя и сохраняет целостность истории поставок.",
    )
    doc.add_picture(str(mocks["sup_index"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 9 – Список поставщиков")

    add_h3(doc, "8.5.2 Создание и редактирование поставщика")
    add_paragraph(
        doc,
        "Маршруты: /Suppliers/Create и /Suppliers/Edit/{id}. Форма включает поля: "
        "«Название» (обязательное), «Телефон», «Адрес». Телефон и адрес необязательны, "
        "но рекомендуются для полноты справочника. Редактирование использует ту же разметку "
        "с предзаполнением значений. Кнопки «Сохранить» и «Отмена» работают стандартно.",
    )
    doc.add_picture(str(mocks["sup_form"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 10 – Форма поставщика")

    # ---- 8.6 Medicines ----
    add_h2(doc, "8.6 Раздел «Препараты»")
    add_h3(doc, "8.6.1 Список препаратов")
    add_paragraph(
        doc,
        "Маршрут: /Medicines. Это основной рабочий экран номенклатуры. Над таблицей расположены:",
    )
    for t in [
        "текстовое поле поиска по названию и производителю;",
        "выпадающий список категорий с пунктом «Все категории»;",
        "кнопки «Найти» и «Сброс»;",
        "кнопка «Создать» для добавления препарата.",
    ]:
        bullet(doc, t)
    add_paragraph(
        doc,
        "Таблица содержит колонки: Название, Производитель, Категория, Цена (с символом ₽), "
        "Остаток, Рецепт (Да/Нет), Действия. Если остаток меньше 10, значение в колонке "
        "«Остаток» выделяется стилем low-stock (красный акцент), что помогает визуально "
        "находить дефицитные позиции даже вне главной страницы.",
    )
    add_paragraph(
        doc,
        "Действие «Изменить» открывает форму редактирования. «Удалить» запрашивает подтверждение "
        "и выполняется только при отсутствии связанных поступлений и позиций продаж.",
    )
    doc.add_picture(str(mocks["med_index"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 11 – Список препаратов")

    add_h3(doc, "8.6.2 Создание и редактирование препарата")
    add_paragraph(
        doc,
        "Маршруты: /Medicines/Create и /Medicines/Edit/{id}. Форма включает следующие элементы управления:",
    )
    for t in [
        "«Название» — обязательная строка;",
        "«Производитель» — обязательная строка;",
        "«Категория» — обязательный выпадающий список, заполняется из таблицы Categories;",
        "«Цена» — числовое поле с шагом 0.01, значение должно быть больше 0;",
        "«Остаток» — целое число ≥ 0 (начальный остаток можно задать вручную, далее он меняется документами);",
        "флажок «Нужен рецепт» — признак рецептурного отпуска.",
    ]:
        bullet(doc, t)
    add_paragraph(
        doc,
        "При сохранении выполняется серверная проверка ModelState. Если категория не выбрана "
        "или числовые поля некорректны, форма отображается повторно с сообщениями "
        "ValidationMessageFor рядом с полями. После успешного создания/обновления "
        "пользователь возвращается к списку препаратов.",
    )
    doc.add_picture(str(mocks["med_form"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 12 – Форма препарата")

    # ---- 8.7 Supplies ----
    add_h2(doc, "8.7 Раздел «Поступления»")
    add_h3(doc, "8.7.1 Журнал поступлений")
    add_paragraph(
        doc,
        "Маршрут: /Supplies. Страница показывает историю приходов товара. Колонки таблицы: "
        "Дата, Препарат, Поставщик, Количество, Действия. Записи упорядочены от новых к старым. "
        "Кнопка «Оформить поступление» открывает форму создания. Для каждой строки доступно "
        "удаление с подтверждением «Удалить поступление и уменьшить остаток?».",
    )
    add_paragraph(
        doc,
        "При удалении система проверяет, что текущий остаток препарата не меньше количества "
        "в удаляемом поступлении. Если проверка не проходит (например, товар уже продан), "
        "удаление блокируется сообщением об ошибке — это защищает от появления отрицательного остатка.",
    )
    doc.add_picture(str(mocks["supply_index"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 13 – Журнал поступлений")

    add_h3(doc, "8.7.2 Оформление поступления")
    add_paragraph(
        doc,
        "Маршрут: /Supplies/Create (опционально ?medicineId=...). Пользователь выбирает препарат "
        "и поставщика из выпадающих списков, указывает количество (минимум 1) и дату поступления. "
        "Дата по умолчанию равна текущему дню. После нажатия «Сохранить» InventoryService "
        "открывает транзакцию, добавляет запись Supply и увеличивает Medicine.StockQuantity "
        "на указанное количество. При успехе показывается сообщение "
        "«Поступление оформлено, остаток увеличен.»",
    )
    add_paragraph(
        doc,
        "Если страница открыта с главной по кнопке «Поступление» у дефицитной позиции, "
        "поле препарата уже будет предвыбрано — достаточно указать поставщика и количество.",
    )
    doc.add_picture(str(mocks["supply_create"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 14 – Оформление поступления")

    # ---- 8.8 Sales ----
    add_h2(doc, "8.8 Раздел «Продажи»")
    add_h3(doc, "8.8.1 Журнал продаж")
    add_paragraph(
        doc,
        "Маршрут: /Sales. Таблица содержит номер продажи, дату, краткий состав (список "
        "«препарат × количество (цена)»), итоговую сумму и действия. Действие «Подробнее» "
        "открывает карточку продажи. «Удалить» возвращает товар на склад: по каждой позиции "
        "остаток увеличивается на проданное количество, затем документ продажи удаляется.",
    )
    doc.add_picture(str(mocks["sale_index"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 15 – Журнал продаж")

    add_h3(doc, "8.8.2 Оформление продажи")
    add_paragraph(
        doc,
        "Маршрут: /Sales/Create. Форма рассчитана на отпуск нескольких позиций за один документ. "
        "Изначально отображаются три строки. В каждой строке есть:",
    )
    for t in [
        "выпадающий список препаратов (в тексте пункта видно текущий остаток и цену);",
        "поле количества;",
        "кнопка «Убрать» для удаления строки.",
    ]:
        bullet(doc, t)
    add_paragraph(
        doc,
        "Кнопка «+ Позиция» добавляет ещё одну строку без перезагрузки страницы. Перед отправкой "
        "рекомендуется заполнить хотя бы одну позицию. Сервер игнорирует пустые строки и "
        "рассчитывает сумму продажи как сумму произведений цены на количество по всем валидным позициям. "
        "Цена фиксируется в SaleItem.UnitPrice на момент продажи, поэтому последующее изменение "
        "цены препарата в справочнике не искажает историю.",
    )
    add_paragraph(
        doc,
        "Если по какой-либо позиции остатка недостаточно, вся операция отменяется, ни одна "
        "позиция не списывается, форма отображается снова с текстом ошибки вида: "
        "«Недостаточно «Название»: в наличии N, запрошено M.»",
    )
    doc.add_picture(str(mocks["sale_create"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 16 – Оформление продажи")

    add_h3(doc, "8.8.3 Карточка продажи")
    add_paragraph(
        doc,
        "Маршрут: /Sales/Details/{id}. Страница показывает номер документа, дату, итоговую сумму "
        "и детальную таблицу позиций: препарат, количество, цена за единицу, сумма по строке. "
        "Кнопка «К списку» возвращает в журнал продаж. Карточка используется для контроля "
        "состава отпуска и сверки суммы.",
    )
    doc.add_picture(str(mocks["sale_details"]), width=Cm(15.5))
    add_caption(doc, "Рисунок 17 – Карточка продажи")

    # ---- 8.9 errors ----
    add_h2(doc, "8.9 Типовые сообщения и ошибки")
    add_paragraph(doc, "Таблица 10 – Сообщения пользователю", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, size=12)
    add_table(doc, ["Ситуация", "Тип", "Текст / поведение"], [
        ["Успешное создание/изменение", "Success", "«… создана/обновлена/удалена»"],
        ["Удаление связанной сущности", "Error", "«Нельзя удалить …: есть связанные …»"],
        ["Недостаточно товара при продаже", "Error/ModelState", "«Недостаточно «…»: в наличии …»"],
        ["Пустая форма продажи", "ModelState", "«Добавьте хотя бы одну позицию…»"],
        ["Дубликат названия справочника", "ModelState", "«Не удалось сохранить. Возможно, такое название уже есть.»"],
        ["Отмена удаления в confirm", "Клиент", "Форма не отправляется"],
        ["Некорректные числовые поля", "Валидация", "Подсветка поля и текст ValidationMessage"],
    ])
    add_paragraph(
        doc,
        "Рекомендуемый порядок работы пользователя в течение дня: проверить главную страницу "
        "и низкие остатки → при необходимости оформить поступления → вести справочники → "
        "оформлять продажи → при ошибочном документе удалить продажу/поступление "
        "(с пониманием влияния на остаток).",
    )

    # ===================== 9 ЗАКЛЮЧЕНИЕ =====================
    add_h1(doc, "9 Заключение")
    add_paragraph(
        doc,
        "В результате выполнения курсовой работы спроектирована и реализована автоматизированная "
        "система «Аптека». Проведён анализ предметной области, сформулированы User Stories и "
        "Use Cases, разработаны схема базы данных, модульная архитектура и схема классов, "
        "выполнен обзор выбранного программного обеспечения.",
    )
    add_paragraph(
        doc,
        "Программная реализация на ASP.NET Core MVC с СУБД PostgreSQL обеспечивает полный "
        "цикл учёта: справочники категорий, поставщиков и препаратов; журнал поступлений; "
        "журнал продаж с многопозиционным отпуском; контроль остатков на главной странице. "
        "Интерфейс построен с использованием HTML Helpers и дополнен JavaScript-сценариями.",
    )
    add_paragraph(
        doc,
        "Поставленные задачи курсовой работы выполнены. Дальнейшее развитие системы может "
        "включать авторизацию пользователей, партионный учёт, отчёты по продажам за период "
        "и интеграцию с устройствами печати чеков.",
    )

    # ===================== ИСТОЧНИКИ =====================
    add_h1(doc, "Список использованных источников")
    sources = [
        "1. Троелсен, Э. Язык программирования C# и платформа .NET / Э. Троелсен, Ф. Джепикс. — М.: Вильямс, 2021. — 1440 с.",
        "2. Фримен, А. ASP.NET Core MVC с примерами на C# для профессионалов / А. Фримен. — М.: Вильямс, 2022.",
        "3. Документация ASP.NET Core MVC [Электронный ресурс]. — URL: https://learn.microsoft.com/aspnet/core/mvc (дата обращения: 11.10.2026).",
        "4. Документация Entity Framework Core [Электронный ресурс]. — URL: https://learn.microsoft.com/ef/core (дата обращения: 11.10.2026).",
        "5. Документация PostgreSQL 17 [Электронный ресурс]. — URL: https://www.postgresql.org/docs/17/ (дата обращения: 11.10.2026).",
        "6. Фаулер, М. UML. Основы. Краткое руководство по стандартному языку объектного моделирования / М. Фаулер. — СПб.: Символ-Плюс, 2019.",
        "7. Коберн, А. Современные методы описания функциональных требований к системам / А. Коберн. — М.: Лори, 2019.",
        "8. Материалы дисциплины «Проектирование и архитектура программных систем»: требования к курсовой работе; пример выполнения курсовой работы.",
        "9. Visual Use Case [Электронный ресурс]. — URL: http://www.visualusecase.com/ (дата обращения: 11.10.2026).",
    ]
    for s in sources:
        add_paragraph(doc, s, first_indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.save(OUT)
    print(f"OK: {OUT}")


if __name__ == "__main__":
    build()
