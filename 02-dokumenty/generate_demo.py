"""Создаёт демо-материалы для генератора документов:
- data/Клиенты.xlsx: таблица клиентов и позиций (её обычно ведёт сам заказчик);
- templates/Договор.docx и templates/Счёт.docx: Word-шаблоны с метками {{ ... }}.

В реальном заказе шаблоны берутся у клиента: в его документ просто вставляются метки.
Все компании и реквизиты в демо вымышленные.
"""
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

BASE = Path(__file__).parent

CLIENTS = [
    ["24-101", "03.11.2026", "ООО «Ромашка»", "7701234567", "генерального директора Иванова Ивана Петровича",
     "Иванов И.П.", "г. Москва, ул. Лесная, д. 5", "info@romashka.example"],
    ["24-102", "05.11.2026", "ИП Смирнова Ольга Викторовна", "500100732259", "",
     "Смирнова О.В.", "г. Казань, пр. Победы, д. 12, оф. 3", "smirnova@example.ru"],
    ["24-103", "10.11.2026", "ООО «Северный ветер»", "7812345670", "директора Кузнецова Дмитрия Алексеевича",
     "Кузнецов Д.А.", "г. Санкт-Петербург, наб. Обводного канала, д. 40", "office@sever.example"],
    ["24-104", "12.11.2026", "ООО «Без позиций»", "123", "директора Тестова Т.Т.",
     "Тестов Т.Т.", "г. Тверь", "test@example.ru"],
]

POSITIONS = [
    ["24-101", "Разработка скрипта автоматизации отчётов", 1, "усл.", 15000],
    ["24-101", "Настройка и обучение сотрудников", 2, "час", 1500],
    ["24-102", "Сборка сводного отчёта по продажам", 3, "мес.", 4000],
    ["24-103", "Парсинг каталога поставщика", 1, "усл.", 8000],
    ["24-103", "Ежедневное обновление данных", 30, "дн.", 250.5],
    ["24-103", "Техническая поддержка", 1, "мес.", 3000],
]


def make_data():
    (BASE / "data").mkdir(exist_ok=True)
    path = BASE / "data" / "Клиенты.xlsx"
    with pd.ExcelWriter(path) as xl:
        pd.DataFrame(CLIENTS, columns=["Номер договора", "Дата", "Компания", "ИНН", "В лице (кого)",
                                       "Подписант", "Адрес", "Email"]).to_excel(xl, sheet_name="Клиенты", index=False)
        pd.DataFrame(POSITIONS, columns=["Номер договора", "Услуга", "Количество", "Ед.",
                                         "Цена"]).to_excel(xl, sheet_name="Позиции", index=False)
    print("Создан", path.relative_to(BASE))


# ---------- Шаблоны ----------

def base_document() -> Document:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    return doc


def para(doc, text, bold=False, align=None, size=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    if align:
        p.alignment = align
    return p


def items_table(doc):
    """Таблица позиций. Строки {%tr ... %} - это цикл docxtpl: строка повторится для каждой услуги."""
    headers = ["№", "Наименование", "Кол-во", "Ед.", "Цена, ₽", "Сумма, ₽"]
    table = doc.add_table(rows=4, cols=6)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        table.cell(0, i).text = h
        table.cell(0, i).paragraphs[0].runs[0].bold = True
    table.cell(1, 0).text = "{%tr for it in items %}"
    values = ["{{ it.n }}", "{{ it.name }}", "{{ it.qty }}", "{{ it.unit }}", "{{ it.price }}", "{{ it.sum }}"]
    for i, v in enumerate(values):
        table.cell(2, i).text = v
    table.cell(3, 0).text = "{%tr endfor %}"
    return table


def make_contract():
    doc = base_document()
    para(doc, "ДОГОВОР ОКАЗАНИЯ УСЛУГ № {{ number }}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    para(doc, "г. Москва\t\t\t\t\t\t\t\t{{ date }}")
    # Для ИП "в лице ..." не пишут, а для ООО пишут; "именуемый/именуемое" считается в скрипте
    para(doc, "{{ seller.name }}, именуемый в дальнейшем «Исполнитель», с одной стороны, и {{ company }}"
              "{% if signer_gen %}, в лице {{ signer_gen }}{% endif %}, {{ called }} в дальнейшем «Заказчик», "
              "с другой стороны, заключили настоящий договор о нижеследующем.")
    para(doc, "1. Предмет договора", bold=True)
    para(doc, "1.1. Исполнитель обязуется оказать, а Заказчик принять и оплатить следующие услуги:")
    items_table(doc)
    para(doc, "2. Стоимость и порядок оплаты", bold=True)
    para(doc, "2.1. Стоимость услуг составляет {{ total }} руб. ({{ total_words }}). НДС не облагается.")
    para(doc, "2.2. Оплата производится в течение 5 (пяти) рабочих дней с даты выставления счёта.")
    para(doc, "3. Реквизиты и подписи сторон", bold=True)
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).text = ("Исполнитель:\n{{ seller.name }}\nИНН {{ seller.inn }}\n{{ seller.address }}\n\n"
                         "_____________ / {{ seller.signer }}")
    t.cell(0, 1).text = ("Заказчик:\n{{ company }}\nИНН {{ inn }}\n{{ address }}\n{{ email }}\n\n"
                         "_____________ / {{ signer }}")
    doc.save(BASE / "templates" / "Договор.docx")


def make_invoice():
    doc = base_document()
    para(doc, "Поставщик: {{ seller.name }}, ИНН {{ seller.inn }}, {{ seller.address }}")
    para(doc, "Банк: {{ seller.bank }}, р/с {{ seller.account }}, БИК {{ seller.bik }}")
    para(doc, "СЧЁТ НА ОПЛАТУ № {{ number }} от {{ date }}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    para(doc, "Покупатель: {{ company }}, ИНН {{ inn }}, {{ address }}")
    para(doc, "Основание: Договор № {{ number }} от {{ date }}")
    items_table(doc)
    para(doc, "Итого: {{ total }} руб.\nБез налога (НДС)", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    para(doc, "Всего наименований {{ items|length }}, на сумму {{ total }} руб.")
    para(doc, "{{ total_words }}", bold=True)
    para(doc, "\nИсполнитель _____________ / {{ seller.signer }}")
    doc.save(BASE / "templates" / "Счёт.docx")


def main():
    make_data()
    (BASE / "templates").mkdir(exist_ok=True)
    make_contract()
    make_invoice()
    print("Созданы шаблоны templates/Договор.docx и templates/Счёт.docx")


if __name__ == "__main__":
    main()
