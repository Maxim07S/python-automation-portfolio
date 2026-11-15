"""Собирает Excel-выгрузки филиалов в один оформленный сводный отчёт.

Что делает скрипт:
1. Читает все .xlsx из папки input/ (один файл = один филиал).
2. Сам понимает разные названия колонок ("Кол-во", "Количество", "кол-во, шт" и т.д.).
3. Чистит данные: пробелы и регистр в названиях, даты текстом, дубли, пустые строки.
4. Строки с ошибками не теряет молча, а выносит на отдельный лист "Проблемы".
5. Сохраняет отчёт: итоги, сводные таблицы, графики, все продажи.

Запуск:
    python merge_reports.py
    python merge_reports.py --input "D:/Отчёты" --output "D:/Сводный.xlsx"
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent

# Все варианты названий колонок, которые встречаются в выгрузках.
# Ключ - вариант (в нижнем регистре, без знаков препинания), значение - стандартное имя.
COLUMN_ALIASES = {
    "дата": "Дата", "дата продажи": "Дата", "дата операции": "Дата",
    "товар": "Товар", "наименование": "Товар", "номенклатура": "Товар",
    "колво": "Количество", "количество": "Количество", "колво шт": "Количество",
    "цена": "Цена", "цена руб": "Цена", "цена за шт": "Цена", "стоимость ед": "Цена",
    "менеджер": "Менеджер", "продавец": "Менеджер", "ответственный": "Менеджер",
}
REQUIRED = ["Дата", "Товар", "Количество", "Цена", "Менеджер"]

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
          "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

RUB = '#,##0" ₽"'


# ---------- Чтение и очистка ----------

def normalize_header(name) -> str:
    text = str(name).lower().replace("ё", "е")
    text = re.sub(r"[^\w\s]", "", text)       # убираем знаки препинания
    return " ".join(text.split())


def parse_date(value):
    """Понимает и настоящие даты Excel, и текст вида 05.07.2026 или 2026-07-05."""
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            pass
    return pd.NaT


def branch_from_file(path: Path) -> str:
    return re.sub(r"^продажи[_\s-]*", "", path.stem, flags=re.IGNORECASE) or path.stem


def read_branch(path: Path, problems: list, stats: dict) -> pd.DataFrame:
    branch = branch_from_file(path)
    raw = pd.read_excel(path)

    rename = {}
    for col in raw.columns:
        std = COLUMN_ALIASES.get(normalize_header(col))
        if std:
            rename[col] = std
    df = raw.rename(columns=rename)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        problems.append({"Файл": path.name, "Строка": "-", "Причина": f"Не найдены колонки: {', '.join(missing)}"})
        return pd.DataFrame()

    df = df[REQUIRED].copy()
    df["Строка в файле"] = df.index + 2          # +2: заголовок и нумерация Excel с 1

    empty = df[REQUIRED].isna().all(axis=1)
    stats["пустых строк"] += int(empty.sum())
    df = df[~empty]

    before = len(df)
    df = df.drop_duplicates(subset=REQUIRED)
    stats["дублей удалено"] += before - len(df)

    df["Дата"] = df["Дата"].map(parse_date)
    df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce")
    df["Цена"] = pd.to_numeric(df["Цена"], errors="coerce")
    df["Товар"] = df["Товар"].astype("string").str.split().str.join(" ")
    df["Менеджер"] = df["Менеджер"].astype("string").str.strip()

    reasons = {
        "Дата": "Не удалось распознать дату",
        "Количество": "Количество не число",
        "Цена": "Цена не число",
        "Товар": "Не указан товар",
    }
    bad_mask = pd.Series(False, index=df.index)
    for col, reason in reasons.items():
        mask = df[col].isna()
        for _, row in df[mask & ~bad_mask].iterrows():
            problems.append({"Файл": path.name, "Строка": int(row["Строка в файле"]), "Причина": reason})
        bad_mask |= mask
    df = df[~bad_mask]

    df.insert(0, "Филиал", branch)
    return df


def unify_product_names(df: pd.DataFrame) -> pd.DataFrame:
    """'КОФЕ ЗЕРНОВОЙ 1 КГ' и 'кофе зерновой 1 кг' -> самое частое написание."""
    key = df["Товар"].str.lower()
    canonical = df.groupby(key)["Товар"].agg(lambda s: s.value_counts().index[0])
    df["Товар"] = key.map(canonical)
    return df


# ---------- Оформление ----------

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_table(ws, money_cols=(), date_cols=(), total_row=False):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = BORDER
            if cell.column in money_cols:
                cell.number_format = RUB
            if cell.column in date_cols:
                cell.number_format = "DD.MM.YYYY"
    if total_row:
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)
            cell.fill = TOTAL_FILL

    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 3, 10), 45)


def add_charts(wb, months_count: int, branches_count: int):
    ws = wb["По филиалам"]
    bar = BarChart()
    bar.title = "Выручка по филиалам"
    bar.y_axis.title = "₽"
    bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=branches_count + 1), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=branches_count + 1))
    bar.legend = None
    bar.height, bar.width = 8, 16
    ws.add_chart(bar, "G2")

    ws = wb["По месяцам"]
    line = LineChart()
    line.title = "Динамика выручки по месяцам"
    line.y_axis.title = "₽"
    line.add_data(Reference(ws, min_col=2, max_col=branches_count + 1, min_row=1, max_row=months_count + 1),
                  titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=2, max_row=months_count + 1))
    line.height, line.width = 9, 18
    ws.add_chart(line, f"{get_column_letter(branches_count + 4)}2")


def write_summary(ws, data: pd.DataFrame, by_branch: pd.DataFrame, files: int, stats: dict, problems: int):
    ws["A1"] = "Сводный отчёт по продажам"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    period = f"{data['Дата'].min():%d.%m.%Y} — {data['Дата'].max():%d.%m.%Y}"
    rows = [
        ("Период", period),
        ("Выручка", data["Сумма"].sum()),
        ("Продаж (строк)", len(data)),
        ("Средний чек", data["Сумма"].mean()),
        ("Лучший филиал", by_branch.iloc[0]["Филиал"]),
        ("Самый продаваемый товар", data.groupby("Товар")["Сумма"].sum().idxmax()),
        ("", ""),
        ("Обработано файлов", files),
        ("Удалено дублей", stats["дублей удалено"]),
        ("Пропущено пустых строк", stats["пустых строк"]),
        ("Строк с ошибками (см. лист «Проблемы»)", problems),
        ("Отчёт сформирован", datetime.now().strftime("%d.%m.%Y %H:%M")),
    ]
    for i, (name, value) in enumerate(rows, start=3):
        ws.cell(i, 1, name).font = Font(bold=bool(name))
        cell = ws.cell(i, 2, value)
        if name in ("Выручка", "Средний чек"):
            cell.number_format = RUB
        cell.alignment = Alignment(horizontal="left")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 28


# ---------- Основная логика ----------

def build_report(input_dir: Path, output: Path):
    files = sorted(p for p in input_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"В папке {input_dir} нет файлов .xlsx")

    problems, stats = [], {"дублей удалено": 0, "пустых строк": 0}
    frames = [read_branch(p, problems, stats) for p in files]
    data = pd.concat([f for f in frames if not f.empty], ignore_index=True)
    data = unify_product_names(data)
    data["Сумма"] = data["Количество"] * data["Цена"]
    data = data.sort_values(["Дата", "Филиал"]).drop(columns="Строка в файле")

    by_branch = (data.groupby("Филиал")
                 .agg(Выручка=("Сумма", "sum"), Продаж=("Сумма", "size"), **{"Средний чек": ("Сумма", "mean")})
                 .sort_values("Выручка", ascending=False).reset_index())
    by_branch["Доля, %"] = (by_branch["Выручка"] / by_branch["Выручка"].sum() * 100).round(1)

    month_key = data["Дата"].dt.to_period("M")
    by_month = data.pivot_table(index=month_key, columns="Филиал", values="Сумма", aggfunc="sum", fill_value=0)
    by_month["Итого"] = by_month.sum(axis=1)
    by_month.index = [f"{MONTHS[p.month - 1]} {p.year}" for p in by_month.index]
    by_month.index.name = "Месяц"
    by_month = by_month.reset_index()

    top = (data.groupby("Товар").agg(Продано_шт=("Количество", "sum"), Выручка=("Сумма", "sum"))
           .sort_values("Выручка", ascending=False).reset_index()
           .rename(columns={"Продано_шт": "Продано, шт"}))

    managers = (data.groupby(["Менеджер", "Филиал"]).agg(Выручка=("Сумма", "sum"), Продаж=("Сумма", "size"))
                .sort_values("Выручка", ascending=False).reset_index())

    def with_total(df, label_col, sum_cols):
        total = {c: "" for c in df.columns}
        total[label_col] = "ИТОГО"
        for c in sum_cols:
            total[c] = df[c].sum()
        return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    problems_df = pd.DataFrame(problems or [{"Файл": "—", "Строка": "—", "Причина": "Ошибок не найдено"}])

    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as xl:
        pd.DataFrame().to_excel(xl, sheet_name="Итоги", index=False)
        with_total(by_branch, "Филиал", ["Выручка", "Продаж"]).to_excel(xl, sheet_name="По филиалам", index=False)
        with_total(by_month, "Месяц", list(by_month.columns[1:])).to_excel(xl, sheet_name="По месяцам", index=False)
        top.to_excel(xl, sheet_name="Топ товаров", index=False)
        managers.to_excel(xl, sheet_name="По менеджерам", index=False)
        data.to_excel(xl, sheet_name="Все продажи", index=False)
        problems_df.to_excel(xl, sheet_name="Проблемы", index=False)

    wb = load_workbook(output)
    write_summary(wb["Итоги"], data, by_branch, len(files), stats, len(problems))
    style_table(wb["По филиалам"], money_cols=(2, 4), total_row=True)
    style_table(wb["По месяцам"], money_cols=range(2, by_month.shape[1] + 1), total_row=True)
    style_table(wb["Топ товаров"], money_cols=(3,))
    style_table(wb["По менеджерам"], money_cols=(3,))
    style_table(wb["Все продажи"], money_cols=(6, 8), date_cols=(2,))
    style_table(wb["Проблемы"])
    add_charts(wb, months_count=len(by_month), branches_count=len(by_branch))
    wb.save(output)

    print(f"Готово: {output}")
    revenue = f"{data['Сумма'].sum():,.0f}".replace(",", " ")
    print(f"  файлов: {len(files)}, продаж: {len(data)}, выручка: {revenue} ₽")
    print(f"  дублей удалено: {stats['дублей удалено']}, пустых строк: {stats['пустых строк']}, "
          f"строк с ошибками: {len(problems)}")


def main():
    parser = argparse.ArgumentParser(description="Сборка Excel-отчётов филиалов в один сводный отчёт")
    parser.add_argument("--input", type=Path, default=BASE / "input", help="папка с файлами филиалов")
    parser.add_argument("--output", type=Path, default=BASE / "output" / "Сводный_отчёт.xlsx")
    args = parser.parse_args()
    build_report(args.input, args.output)


if __name__ == "__main__":
    main()
