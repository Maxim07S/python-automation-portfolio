"""Создаёт демонстрационные Excel-файлы филиалов в папке input/.

Файлы специально "грязные", как в реальной жизни:
- в разных филиалах колонки называются по-разному и стоят в разном порядке;
- лишние пробелы и разный регистр в названиях товаров;
- даты то датой, то текстом;
- дубли строк, пустые строки, строки с ошибками.
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)

BASE = Path(__file__).parent
INPUT = BASE / "input"

PRODUCTS = {
    "Кофе зерновой 1 кг": 1450,
    "Кофе молотый 250 г": 520,
    "Чай чёрный 100 г": 310,
    "Чай зелёный 100 г": 340,
    "Сироп ванильный": 590,
    "Капсулы 10 шт": 450,
    "Кружка фирменная": 690,
    "Турка медная": 2300,
}

# У каждого филиала свой "формат" выгрузки
BRANCHES = {
    "Москва": {"Дата": "Дата", "Товар": "Товар", "Кол-во": "Кол-во", "Цена": "Цена", "Менеджер": "Менеджер"},
    "Санкт-Петербург": {"Дата": "Дата продажи", "Товар": "Наименование", "Кол-во": "Количество", "Цена": "Цена, руб.", "Менеджер": "Продавец"},
    "Казань": {"Дата": "дата", "Товар": "товар", "Кол-во": "кол-во, шт", "Цена": "цена за шт", "Менеджер": "менеджер"},
    "Екатеринбург": {"Дата": "Дата", "Товар": "Номенклатура", "Кол-во": "Кол-во", "Цена": "Цена", "Менеджер": "Ответственный"},
    "Новосибирск": {"Дата": "Дата операции", "Товар": "Товар", "Кол-во": "Количество", "Цена": "Стоимость ед.", "Менеджер": "Менеджер"},
}

MANAGERS = ["Иванова А.", "Петров С.", "Сидорова Е.", "Кузнецов Д.", "Смирнова О.", "Волков И."]


def messy_name(name: str) -> str:
    """Портит название товара так, как это делают люди."""
    variant = random.random()
    if variant < 0.15:
        return "  " + name + " "
    if variant < 0.25:
        return name.upper()
    if variant < 0.32:
        return name.lower()
    return name


def messy_date(d: date):
    """Иногда дата записана текстом в разных форматах."""
    variant = random.random()
    if variant < 0.2:
        return d.strftime("%d.%m.%Y")
    if variant < 0.3:
        return d.strftime("%Y-%m-%d")
    return pd.Timestamp(d)


def make_branch(branch: str, columns: dict) -> pd.DataFrame:
    rows = []
    start = date(2026, 7, 1)
    managers = random.sample(MANAGERS, 2)
    for _ in range(random.randint(120, 200)):
        product = random.choice(list(PRODUCTS))
        d = start + timedelta(days=random.randint(0, 91))
        rows.append({
            columns["Дата"]: messy_date(d),
            columns["Товар"]: messy_name(product),
            columns["Кол-во"]: random.choice([1, 1, 1, 2, 2, 3, 5]),
            columns["Цена"]: PRODUCTS[product],
            columns["Менеджер"]: random.choice(managers),
        })

    df = pd.DataFrame(rows)
    # Дубли (кто-то дважды вставил одни и те же строки)
    df = pd.concat([df, df.sample(5, random_state=1)], ignore_index=True).astype(object)
    # Пустая строка и строка с ошибкой в количестве
    df.loc[len(df)] = [None] * len(df.columns)
    bad = df.iloc[0].copy()
    bad[columns["Кол-во"]] = "два"
    df.loc[len(df)] = bad
    # Перемешиваем порядок колонок, как в реальных выгрузках
    cols = list(df.columns)
    random.shuffle(cols)
    return df[cols].sample(frac=1, random_state=2).reset_index(drop=True)


def main():
    INPUT.mkdir(exist_ok=True)
    for branch, columns in BRANCHES.items():
        path = INPUT / f"Продажи_{branch}.xlsx"
        make_branch(branch, columns).to_excel(path, index=False)
        print("Создан", path.name)


if __name__ == "__main__":
    main()
