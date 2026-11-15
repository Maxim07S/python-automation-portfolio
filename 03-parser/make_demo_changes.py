"""Только для демонстрации: на учебном сайте цены не меняются, поэтому этот скрипт
создаёт "вчерашний" снимок с изменёнными ценами. После него запустите parse_catalog.py
ещё раз, и на листе «Изменения» появятся подорожания, скидки, новые и пропавшие товары.
"""
from pathlib import Path

import pandas as pd

history = Path(__file__).parent / "history"
latest = sorted(history.glob("snapshot_*.csv"))[-1]
df = pd.read_csv(latest)

sample = df.sample(15, random_state=7).index
df.loc[sample[:8], "Цена"] = (df.loc[sample[:8], "Цена"] * 0.85).round(2)    # сейчас будут "подорожавшими"
df.loc[sample[8:], "Цена"] = (df.loc[sample[8:], "Цена"] * 1.20).round(2)    # сейчас будут "подешевевшими"
df = df.drop(df.sample(3, random_state=1).index)                             # сейчас будут "новыми"
gone = pd.DataFrame([
    {"Название": "Снятая с продажи книга 1", "Категория": "Travel", "Цена": 19.99, "Рейтинг": 3,
     "В наличии": "Да", "Ссылка": "https://books.toscrape.com/catalogue/removed-1/index.html"},
    {"Название": "Снятая с продажи книга 2", "Категория": "Poetry", "Цена": 7.50, "Рейтинг": 1,
     "В наличии": "Да", "Ссылка": "https://books.toscrape.com/catalogue/removed-2/index.html"},
])
df = pd.concat([df, gone], ignore_index=True)

out = latest.with_name(latest.stem + "_demo.csv")
df.to_csv(out, index=False, encoding="utf-8-sig")
print("Создан демо-снимок:", out.name)
