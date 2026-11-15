"""Сумма прописью для документов: 12345.50 -> 'Двенадцать тысяч триста сорок пять рублей 50 копеек'."""

UNITS_M = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
UNITS_F = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
TEENS = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать",
         "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят",
        "восемьдесят", "девяносто"]
HUNDREDS = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот",
            "восемьсот", "девятьсот"]

# (формы слова для 1, 2-4, 5+), род
SCALES = [
    (("", "", ""), "m"),
    (("тысяча", "тысячи", "тысяч"), "f"),
    (("миллион", "миллиона", "миллионов"), "m"),
    (("миллиард", "миллиарда", "миллиардов"), "m"),
]


def plural(n: int, forms: tuple) -> str:
    """Выбирает форму слова: 1 рубль, 2 рубля, 5 рублей, 11 рублей, 21 рубль."""
    n = abs(n) % 100
    if 11 <= n <= 19:
        return forms[2]
    n %= 10
    if n == 1:
        return forms[0]
    if 2 <= n <= 4:
        return forms[1]
    return forms[2]


def triad_to_words(n: int, gender: str) -> list:
    words = [HUNDREDS[n // 100]]
    rest = n % 100
    if 10 <= rest <= 19:
        words.append(TEENS[rest - 10])
    else:
        words.append(TENS[rest // 10])
        words.append((UNITS_F if gender == "f" else UNITS_M)[rest % 10])
    return [w for w in words if w]


def number_to_words(n: int) -> str:
    if n == 0:
        return "ноль"
    parts, scale = [], 0
    while n > 0:
        triad = n % 1000
        if triad:
            forms, gender = SCALES[scale]
            chunk = triad_to_words(triad, gender)
            if scale:
                chunk.append(plural(triad, forms))
            parts.insert(0, " ".join(chunk))
        n //= 1000
        scale += 1
    return " ".join(parts)


def rubles_in_words(amount: float) -> str:
    kopecks_total = round(amount * 100)
    rubles, kopecks = divmod(kopecks_total, 100)
    text = number_to_words(rubles)
    rub = plural(rubles, ("рубль", "рубля", "рублей"))
    kop = plural(kopecks, ("копейка", "копейки", "копеек"))
    return f"{text[0].upper()}{text[1:]} {rub} {kopecks:02d} {kop}"


if __name__ == "__main__":
    for x in [0, 1, 2, 5, 11, 21, 101, 1000, 1001, 2000, 5000, 21000, 112, 1234567.89, 45000.5]:
        print(f"{x:>12} -> {rubles_in_words(x)}")
