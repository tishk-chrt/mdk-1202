"""
Анализ цифрового следа и репутации компании в интернете.

Скрипт автоматизирует ручную часть практической работы:
  1. читает CSV с упоминаниями компаний (площадка, URL, рейтинг, текст, ответ компании);
  2. определяет тональность упоминания (словарный метод + числовой рейтинг);
  3. выделяет главные темы жалоб и похвалы (качество, цена, доставка, обман и т.д.);
  4. оценивает качество реакции компании (игнор / шаблонная отписка / конструктив);
  5. считает индекс репутации 0–100 и уровень репутационного риска;
  6. сохраняет итоговую таблицу в Excel (лист «Упоминания» + лист «Сводка»)
     и краткий отчёт в Markdown, печатает сводку в консоль.

Запуск:
    python reputation_analyzer.py                       # mentions.csv -> output/
    python reputation_analyzer.py -i data.csv -o result

Зависимости: только стандартная библиотека; для Excel — openpyxl (pip install openpyxl).
"""

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

# --- Словари для тонального анализа (основы слов, чтобы ловить разные окончания) ---

POSITIVE = [
    "хвал", "удобн", "быстр", "вкусн", "вежлив", "отличн", "хорош", "нрав", "люб",
    "рекоменд", "свеж", "горяч", "прозрачн", "выгодн", "широк", "низк", "гибк",
    "своевременн", "интересн", "оперативн", "довол", "спасибо", "без проблем",
]
NEGATIVE = [
    "ругают", "жалоб", "брак", "подделк", "задерж", "долг", "холодн", "потер", "повреж",
    "обман", "мошен", "блокир", "навяз", "штраф", "хаос", "переработ", "сбой", "сбои", "ошибк",
    "не приехал", "не дозвон", "невозможн", "худш", "ухудш", "сниж", "уменьш",
    "не соответств", "несоответств", "безответствен", "присвоил", "дорог", "платн",
    "очеред", "срыв", "бот", "отмен", "занижен", "нарушен", "без ответа", "без объяснен",
]

# Темы упоминаний: за что хвалят / ругают
TOPICS = {
    "Качество товара/услуги": ["качеств", "брак", "подделк", "ингредиент", "начинк", "вкус", "тесто", "свеж"],
    "Цена и тарифы": ["цен", "стоимост", "дорог", "тариф", "кешбэк", "ставк", "акци"],
    "Доставка и сроки": ["доставк", "срок", "задерж", "посылк", "подач", "не приехал", "курьер"],
    "Сервис и поддержка": ["поддержк", "горяч", "чат", "бот", "дозвон", "сервис", "сотрудник"],
    "Обман и блокировки": ["обман", "мошен", "блокир", "присвоил", "списан", "навяз"],
    "Приложение и сайт": ["приложен", "уведомлен", "поиск", "сайт", "суперприложен"],
    "Условия труда (HR)": ["график", "зарплат", "штраф", "нагрузк", "оплат", "переработ"],
    "Возврат и компенсации": ["возврат", "компенсац", "вернул"],
}

# Маркеры качества ответа компании
TEMPLATE_PHRASES = [
    "спасибо за отзыв", "спасибо за обратную связь", "спасибо за вашу обратную связь",
    "благодарим за", "нам жаль", "приносим извинения", "всегда на связи", "с уважением",
    "мы высоко ценим", "мы ценим",
]
CONSTRUCTIVE_MARKERS = [
    "номер заказа", "номер обращения", "номер договора", "личн", "вернули", "вернём",
    "компенсац", "возврат", "разбира", "разберёмся", "проверим", "передали", "передаём",
    "исправ", "решен",
]

REPLY_SCORE = {"Конструктивное решение": 100, "Шаблонная отписка": 40, "Игнорирует": 0}


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").lower().replace("ё", "е"))


def count_hits(text: str, stems: list[str]) -> int:
    """Считает слова, которые НАЧИНАЮТСЯ с основы (чтобы «бот» не находился в «работе»)."""
    return sum(len(re.findall(r"(?<![а-яa-z])" + re.escape(normalize(s)), text)) for s in stems)


def text_sentiment(text: str) -> float:
    """Оценка тональности текста от -1 (негатив) до +1 (позитив)."""
    t = normalize(text)
    pos, neg = count_hits(t, POSITIVE), count_hits(t, NEGATIVE)
    # «не удобно», «не быстро» — отрицание переворачивает позитивное слово
    flipped = len(re.findall(r"\bне (?:" + "|".join(POSITIVE[:12]) + r")", t))
    pos, neg = pos - flipped, neg + flipped
    if pos + neg == 0:
        return 0.0
    return (pos - neg) / (pos + neg)


def rating_to_5(rating: str, scale: str) -> float | None:
    """Приводит рейтинг к пятибалльной шкале (например, 86,12 из 100 -> 4,31)."""
    try:
        r, s = float(rating.replace(",", ".")), float(scale.replace(",", "."))
    except (ValueError, AttributeError):
        return None
    return round(r / s * 5, 2) if s else None


def classify_tone(text: str, rating5: float | None) -> tuple[str, float]:
    """Итоговая тональность: числовой рейтинг (вес 0,6) + анализ текста (вес 0,4)."""
    score = text_sentiment(text)
    if rating5 is not None:
        rating_score = (rating5 - 3) / 2          # 1 -> -1, 3 -> 0, 5 -> +1
        score = 0.6 * rating_score + 0.4 * score
    if score > 0.15:
        return "Позитивная", score
    if score < -0.15:
        return "Негативная", score
    return "Нейтральная", score


def detect_topics(text: str) -> list[str]:
    t = normalize(text)
    hits = {name: count_hits(t, stems) for name, stems in TOPICS.items()}
    return [name for name, n in sorted(hits.items(), key=lambda x: -x[1]) if n > 0][:3]


def classify_reply(reply: str, mention: str) -> str:
    r = normalize(reply)
    if r in ("", "-", "—"):
        return "Игнорирует"
    if r == "н/д":
        return "Не применимо"
    if any(m in r for m in CONSTRUCTIVE_MARKERS):
        return "Конструктивное решение"
    # Персональный ответ: компания пересказывает суть жалобы, а не пишет общие слова
    words = {w[:6] for w in re.findall(r"[а-я]{6,}", r)}
    mention_words = {w[:6] for w in re.findall(r"[а-я]{6,}", normalize(mention))}
    template_hits = count_hits(r, TEMPLATE_PHRASES)
    if len(words & mention_words) >= 2 and template_hits == 0:
        return "Конструктивное решение"
    return "Шаблонная отписка"


def risk_level(index: float) -> str:
    if index < 40:
        return "Высокий"
    if index < 60:
        return "Средний"
    return "Низкий"


def analyze(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    mentions = []
    for i, row in enumerate(rows, 1):
        rating5 = rating_to_5(row.get("rating", ""), row.get("rating_scale", ""))
        tone, score = classify_tone(row["text"], rating5)
        rating_view = "—"
        if rating5 is not None:
            rating_view = f'{row["rating"]} из {row["rating_scale"]}'.replace(".", ",")
        mentions.append({
            "№": i,
            "Компания": row["company"],
            "URL / Площадка": f'{row["platform"]}\n{row["url"]}',
            "Тип ресурса": row["resource_type"],
            "Оценка / Рейтинг": rating_view,
            "Рейтинг (из 5)": rating5,
            "Тональность": tone,
            "Балл тональности": round(score, 2),
            "Главные темы": ", ".join(detect_topics(row["text"])) or "—",
            "Суть упоминаний / Главные проблемы": row["text"],
            "Качество реакции компании": classify_reply(row.get("company_reply", ""), row["text"]),
        })

    by_company = defaultdict(list)
    for m in mentions:
        by_company[m["Компания"]].append(m)

    summary = []
    for company, items in by_company.items():
        n = len(items)
        tones = Counter(m["Тональность"] for m in items)
        ratings = [m["Рейтинг (из 5)"] for m in items if m["Рейтинг (из 5)"] is not None]
        avg5 = sum(ratings) / len(ratings) if ratings else 3.0
        replies = [m["Качество реакции компании"] for m in items
                   if m["Качество реакции компании"] != "Не применимо"]
        reply_score = sum(REPLY_SCORE[r] for r in replies) / len(replies) if replies else 0
        topics = Counter(t for m in items for t in m["Главные темы"].split(", ") if t != "—")
        # Индекс репутации: 50 % — средний рейтинг, 30 % — доля позитива, 20 % — работа с отзывами
        index = 0.5 * (avg5 / 5 * 100) + 0.3 * (tones["Позитивная"] / n * 100) + 0.2 * reply_score
        summary.append({
            "Компания": company,
            "Упоминаний": n,
            "Средний рейтинг (из 5)": round(avg5, 2),
            "Позитивных": tones["Позитивная"],
            "Нейтральных": tones["Нейтральная"],
            "Негативных": tones["Негативная"],
            "Преобладающая тональность": tones.most_common(1)[0][0],
            "Главные темы": ", ".join(t for t, _ in topics.most_common(3)),
            "Ответы: конструктив / шаблон / игнор": "{} / {} / {}".format(
                replies.count("Конструктивное решение"),
                replies.count("Шаблонная отписка"),
                replies.count("Игнорирует")),
            "Индекс репутации (0–100)": round(index, 1),
            "Репутационный риск": risk_level(index),
        })
    summary.sort(key=lambda s: -s["Индекс репутации (0–100)"])
    return mentions, summary


def save_excel(mentions: list[dict], summary: list[dict], path: Path) -> bool:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        return False

    fills = {"Позитивная": "C6EFCE", "Негативная": "FFC7CE", "Нейтральная": "FFEB9C",
             "Высокий": "FFC7CE", "Средний": "FFEB9C", "Низкий": "C6EFCE"}
    wb = Workbook()
    for ws, data in ((wb.active, mentions), (wb.create_sheet(), summary)):
        ws.title = "Упоминания" if data is mentions else "Сводка"
        headers = list(data[0].keys())
        ws.append(headers)
        for row in data:
            ws.append([row[h] for h in headers])
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
        for col in ws.iter_cols(min_row=1):
            for cell in col:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if cell.row > 1 and cell.value in fills:
                    cell.fill = PatternFill("solid", fgColor=fills[cell.value])
            width = max(len(str(c.value or "").split("\n")[0]) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(width, 10), 60)
        ws.freeze_panes = "A2"
    wb.save(path)
    return True


def save_markdown(summary: list[dict], path: Path) -> None:
    lines = ["# Сводка по репутации компаний", ""]
    headers = list(summary[0].keys())
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("|" + "---|" * len(headers))
    for s in summary:
        lines.append("| " + " | ".join(str(s[h]) for h in headers) + " |")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def print_console(summary: list[dict]) -> None:
    print("\nИндекс репутации (0–100)")
    print("-" * 60)
    for s in summary:
        idx = s["Индекс репутации (0–100)"]
        bar = "█" * int(idx // 4)
        print(f'{s["Компания"]:<12} {bar:<25} {idx:>5}  риск: {s["Репутационный риск"]}')
    print("-" * 60)


def main() -> None:
    base = Path(__file__).parent
    parser = argparse.ArgumentParser(description="Анализ цифрового следа и репутации компаний")
    parser.add_argument("-i", "--input", default=base / "mentions.csv", type=Path)
    parser.add_argument("-o", "--output", default=base / "output", type=Path)
    args = parser.parse_args()

    with open(args.input, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))

    mentions, summary = analyze(rows)
    args.output.mkdir(parents=True, exist_ok=True)

    (args.output / "result.json").write_text(
        json.dumps({"mentions": mentions, "summary": summary}, ensure_ascii=False, indent=2),
        encoding="utf-8")
    save_markdown(summary, args.output / "summary.md")
    excel_ok = save_excel(mentions, summary, args.output / "reputation_table.xlsx")

    print(f"Обработано упоминаний: {len(mentions)}, компаний: {len(summary)}")
    print_console(summary)
    print(f"Результаты сохранены в: {args.output}")
    if not excel_ok:
        print("Excel не создан: установите openpyxl (pip install openpyxl)")


if __name__ == "__main__":
    main()
