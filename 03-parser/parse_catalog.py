"""Парсер каталога интернет-магазина с выгрузкой в Excel и отслеживанием изменений цен.

Демо-сайт: https://books.toscrape.com (создан специально для практики парсинга).

Что делает:
1. Обходит все категории и все страницы каталога (вежливо: с паузой и повторами при ошибках).
2. Собирает название, категорию, цену, рейтинг, наличие и ссылку.
3. Сохраняет снимок в history/ (CSV с датой и временем).
4. Сравнивает с прошлым снимком: новые товары, пропавшие товары, изменения цен.
5. Делает оформленный Excel: каталог, статистика по категориям, изменения.

Запуск:
    python parse_catalog.py
    python parse_catalog.py --limit 5     # только первые 5 категорий (для быстрой проверки)
"""
import argparse
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
START_URL = "https://books.toscrape.com/index.html"
DELAY = 0.3             # пауза между запросами, чтобы не нагружать сайт
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
PRICE_FMT = '"£"#,##0.00'


class Fetcher:
    """HTTP-клиент с паузой между запросами и повторами при сбоях."""

    def __init__(self, delay=DELAY, retries=3):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "Mozilla/5.0 (catalog-parser demo)"
        self.delay, self.retries = delay, retries
        self.requests_made = 0

    def get(self, url: str) -> BeautifulSoup:
        for attempt in range(1, self.retries + 1):
            try:
                time.sleep(self.delay)
                resp = self.session.get(url, timeout=20)
                resp.raise_for_status()
                self.requests_made += 1
                return BeautifulSoup(resp.content, "html.parser")
            except requests.RequestException as e:
                if attempt == self.retries:
                    raise
                print(f"    ! {e}. Повтор {attempt}/{self.retries - 1}...")
                time.sleep(2 * attempt)


def get_categories(fetcher: Fetcher) -> list:
    soup = fetcher.get(START_URL)
    links = soup.select("div.side_categories ul li ul li a")
    return [(a.get_text(strip=True), urljoin(START_URL, a["href"])) for a in links]


def parse_product(card, page_url: str, category: str) -> dict:
    link = card.select_one("h3 a")
    price_text = card.select_one("p.price_color").get_text(strip=True)
    rating_class = [c for c in card.select_one("p.star-rating")["class"] if c != "star-rating"]
    return {
        "Название": link["title"],
        "Категория": category,
        "Цена": float(re.sub(r"[^\d.]", "", price_text)),
        "Рейтинг": RATINGS.get(rating_class[0] if rating_class else "", None),
        "В наличии": "Да" if "In stock" in card.select_one("p.availability").get_text() else "Нет",
        "Ссылка": urljoin(page_url, link["href"]),
    }


def parse_category(fetcher: Fetcher, name: str, url: str) -> list:
    products = []
    while url:
        soup = fetcher.get(url)
        products += [parse_product(card, url, name) for card in soup.select("article.product_pod")]
        nxt = soup.select_one("ul.pager li.next a")
        url = urljoin(url, nxt["href"]) if nxt else None
    return products


# ---------- Сравнение с прошлым запуском ----------

def previous_snapshot(history: Path, current: Path):
    files = sorted(p for p in history.glob("snapshot_*.csv") if p != current)
    return pd.read_csv(files[-1]) if files else None, (files[-1].name if files else None)


def compare(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """Сравнивает по ссылке на товар: она уникальна, в отличие от названия."""
    merged = old.merge(new, on="Ссылка", how="outer", suffixes=(" было", " стало"), indicator=True)
    rows = []
    for _, r in merged.iterrows():
        if r["_merge"] == "right_only":
            rows.append(("Новый товар", r["Название стало"], r["Категория стало"], None, r["Цена стало"]))
        elif r["_merge"] == "left_only":
            rows.append(("Пропал из каталога", r["Название было"], r["Категория было"], r["Цена было"], None))
        elif abs(r["Цена было"] - r["Цена стало"]) > 0.001:
            kind = "Подорожал" if r["Цена стало"] > r["Цена было"] else "Подешевел"
            rows.append((kind, r["Название стало"], r["Категория стало"], r["Цена было"], r["Цена стало"]))
    df = pd.DataFrame(rows, columns=["Изменение", "Название", "Категория", "Цена была", "Цена стала"])
    df["Разница, %"] = ((df["Цена стала"] / df["Цена была"] - 1) * 100).round(1)
    return df.sort_values(["Изменение", "Разница, %"]).reset_index(drop=True)


# ---------- Excel ----------

def style_sheet(ws, price_cols=()):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="375623")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col[:300])
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 60)
    for col in price_cols:
        for cell in ws[get_column_letter(col)][1:]:
            cell.number_format = PRICE_FMT


def save_excel(df: pd.DataFrame, changes, prev_name, path: Path):
    by_cat = (df.groupby("Категория")
              .agg(Товаров=("Цена", "size"), **{"Средняя цена": ("Цена", "mean"), "Мин. цена": ("Цена", "min"),
                                                 "Макс. цена": ("Цена", "max"), "Средний рейтинг": ("Рейтинг", "mean")})
              .round(2).sort_values("Товаров", ascending=False).reset_index())

    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Каталог", index=False)
        by_cat.to_excel(xl, sheet_name="По категориям", index=False)
        if changes is None:
            pd.DataFrame({"Изменение": ["Первый запуск: сравнивать пока не с чем"]}).to_excel(
                xl, sheet_name="Изменения", index=False)
        else:
            changes.to_excel(xl, sheet_name="Изменения", index=False)

    wb = load_workbook(path)
    ws = wb["Каталог"]
    style_sheet(ws, price_cols=(3,))
    for cell in ws["F"][1:]:                       # кликабельные ссылки
        cell.hyperlink = cell.value
        cell.style = "Hyperlink"
    style_sheet(wb["По категориям"], price_cols=(3, 4, 5))
    ws = wb["Изменения"]
    style_sheet(ws, price_cols=(4, 5))
    colors = {"Подорожал": "F8CBAD", "Подешевел": "C6EFCE", "Новый товар": "DDEBF7", "Пропал из каталога": "D9D9D9"}
    for row in ws.iter_rows(min_row=2):
        fill = colors.get(row[0].value)
        if fill:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=fill)
    if prev_name:
        ws.cell(1, 8, f"Сравнение с: {prev_name}").font = Font(italic=True)
    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description="Парсер каталога с отслеживанием цен")
    parser.add_argument("--limit", type=int, help="обработать только N первых категорий")
    args = parser.parse_args()

    started = datetime.now()
    fetcher = Fetcher()
    categories = get_categories(fetcher)[: args.limit]
    print(f"Категорий: {len(categories)}")

    products = []
    for i, (name, url) in enumerate(categories, 1):
        items = parse_category(fetcher, name, url)
        products += items
        print(f"  [{i}/{len(categories)}] {name}: {len(items)}")

    df = pd.DataFrame(products).drop_duplicates("Ссылка")
    history = BASE / "history"
    history.mkdir(exist_ok=True)
    snapshot = history / f"snapshot_{started:%Y-%m-%d_%H-%M}.csv"
    df.to_csv(snapshot, index=False, encoding="utf-8-sig")

    old, prev_name = previous_snapshot(history, snapshot)
    changes = compare(old, df) if old is not None else None

    out = BASE / "output"
    out.mkdir(exist_ok=True)
    xlsx = out / f"Каталог_{started:%Y-%m-%d}.xlsx"
    save_excel(df, changes, prev_name, xlsx)

    took = (datetime.now() - started).total_seconds()
    print(f"\nГотово за {took:.0f} с: {len(df)} товаров, {fetcher.requests_made} запросов → {xlsx}")
    if changes is not None:
        print("Изменения с прошлого запуска:", changes["Изменение"].value_counts().to_dict() or "нет")


if __name__ == "__main__":
    main()
