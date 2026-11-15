"""Генерирует договоры и счета в Word по таблице клиентов.

Вход:
    data/Клиенты.xlsx: лист "Клиенты" (одна строка = один договор) и лист "Позиции" (услуги);
    templates/*.docx: Word-шаблоны с метками {{ company }}, {{ total }} и т.д.
Выход:
    output/<номер>_<компания>/Договор.docx, Счёт.docx
    output/Реестр.xlsx: список сформированных документов и строки с ошибками.

Запуск:
    python make_documents.py
    python make_documents.py --only 24-101        # только один договор
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

from rub_words import rubles_in_words

BASE = Path(__file__).parent

# Реквизиты исполнителя (того, кто выставляет документы). Вымышленные.
SELLER = {
    "name": "ИП Петров Сергей Андреевич",
    "signer": "Петров С.А.",
    "inn": "771234567890",
    "address": "г. Москва, ул. Примерная, д. 1",
    "bank": "АО «Демо-Банк»",
    "account": "40802810000000000001",
    "bik": "044500000",
}


def money(x: float) -> str:
    """15000.5 -> '15 000,50' (как принято в документах)."""
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def safe_filename(text: str) -> str:
    text = re.sub(r'[«»"<>:/\\|?*]', "", text)
    return " ".join(text.split())[:60]


def is_ip(company: str) -> bool:
    return str(company).strip().upper().startswith(("ИП ", "ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ"))


def check_client(row) -> list:
    errors = []
    inn = str(row["ИНН"]).strip()
    if not re.fullmatch(r"\d{10}|\d{12}", inn):
        errors.append(f"ИНН «{inn}» должен состоять из 10 или 12 цифр")
    for col in ["Номер договора", "Дата", "Компания", "Подписант"]:
        if not str(row[col]).strip():
            errors.append(f"Не заполнено поле «{col}»")
    if not is_ip(row["Компания"]) and not str(row["В лице (кого)"]).strip():
        errors.append("Для организации нужно заполнить «В лице (кого)»")
    return errors


def build_context(row, positions: pd.DataFrame) -> dict:
    items, total = [], 0.0
    for n, (_, p) in enumerate(positions.iterrows(), start=1):
        line_sum = round(float(p["Количество"]) * float(p["Цена"]), 2)
        total += line_sum
        items.append({
            "n": n, "name": p["Услуга"], "qty": f"{p['Количество']:g}", "unit": p["Ед."],
            "price": money(float(p["Цена"])), "sum": money(line_sum),
        })
    return {
        "number": row["Номер договора"], "date": row["Дата"],
        "company": row["Компания"], "inn": row["ИНН"], "address": row["Адрес"], "email": row["Email"],
        "signer": row["Подписант"],
        "signer_gen": "" if is_ip(row["Компания"]) else row["В лице (кого)"],
        "called": "именуемый" if is_ip(row["Компания"]) else "именуемое",
        "items": items, "total": money(total), "total_words": rubles_in_words(total),
        "seller": SELLER, "_total_num": total,
    }


def main():
    parser = argparse.ArgumentParser(description="Договоры и счета в Word из таблицы Excel")
    parser.add_argument("--data", type=Path, default=BASE / "data" / "Клиенты.xlsx")
    parser.add_argument("--templates", type=Path, default=BASE / "templates")
    parser.add_argument("--output", type=Path, default=BASE / "output")
    parser.add_argument("--only", help="номер договора, если нужен только один")
    args = parser.parse_args()

    # dtype=str, чтобы ИНН с нулём в начале не превратился в число
    clients = pd.read_excel(args.data, sheet_name="Клиенты", dtype=str).fillna("")
    positions = pd.read_excel(args.data, sheet_name="Позиции", dtype={"Номер договора": str})
    templates = sorted(args.templates.glob("*.docx"))
    if not templates:
        raise SystemExit(f"Нет шаблонов .docx в {args.templates}")
    if args.only:
        clients = clients[clients["Номер договора"] == args.only]

    registry = []
    for _, row in clients.iterrows():
        number = row["Номер договора"]
        errors = check_client(row)
        own = positions[positions["Номер договора"] == number]
        if own.empty:
            errors.append("Нет ни одной позиции на листе «Позиции»")
        if errors:
            registry.append({"Договор": number, "Компания": row["Компания"], "Сумма": None,
                             "Статус": "ОШИБКА: " + "; ".join(errors), "Папка": ""})
            print(f"  ✗ {number}: {'; '.join(errors)}")
            continue

        context = build_context(row, own)
        folder = args.output / safe_filename(f"{number}_{row['Компания']}")
        folder.mkdir(parents=True, exist_ok=True)
        for template in templates:
            doc = DocxTemplate(template)
            doc.render(context)
            doc.save(folder / template.name)
        registry.append({"Договор": number, "Компания": row["Компания"], "Сумма": context["_total_num"],
                         "Статус": "Готово", "Папка": str(folder.relative_to(args.output))})
        print(f"  ✓ {number} {row['Компания']}: {context['total']} ₽ → {folder.name}")

    args.output.mkdir(parents=True, exist_ok=True)
    reg = pd.DataFrame(registry)
    reg_path = args.output / "Реестр.xlsx"
    with pd.ExcelWriter(reg_path, engine="openpyxl") as xl:
        reg.to_excel(xl, sheet_name="Реестр", index=False)
        ws = xl.sheets["Реестр"]
        for col, width in zip("ABCDE", (12, 32, 14, 60, 40)):
            ws.column_dimensions[col].width = width
        for cell in ws["C"][1:]:
            cell.number_format = '#,##0.00" ₽"'
        ws["G1"] = f"Сформировано {datetime.now():%d.%m.%Y %H:%M}"

    ok = (reg["Статус"] == "Готово").sum()
    print(f"Готово: {ok} из {len(reg)} договоров. Реестр: {reg_path}")


if __name__ == "__main__":
    main()
