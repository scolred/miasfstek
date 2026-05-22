#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматическое создание еженедельных отчётов МИАЦ и ФСТЭК.

Что делает:
  - Берёт два шаблона .docx (МИАЦ и ФСТЭК).
  - Находит в них строку вида "за <число> <месяц> инцидентов не зарегистрировано".
  - Создаёт по одному файлу на каждый день недели (Пн–Вс) с подставленной датой.
  - Сохраняет файлы с именами вида "МИАЦ ДДММГГ.docx" и "ФСТЭК ДДММГГ.docx".
  - Месяц меняется автоматически (если неделя переходит на новый — фраза станет
    "за 1 июня инцидентов не зарегистрировано" и т.п.).

Как пользоваться:
  python generate_reports.py                  # текущая неделя (с понедельника)
  python generate_reports.py 18.05.2026       # неделя, начинающаяся с указанной даты
  python generate_reports.py --next           # следующая неделя относительно сегодня

Пути к шаблонам и папке вывода задаются в блоке НАСТРОЙКИ ниже.
Внешние зависимости не нужны — только стандартная библиотека Python 3.8+.
"""

import argparse
import re
import sys
import zipfile
from datetime import date, timedelta
from pathlib import Path

# ============================================================================
# НАСТРОЙКИ — поменяй пути под свои файлы
# ============================================================================

# Шаблоны лежат рядом со скриптом. Подойдёт любой готовый отчёт за прошлую неделю —
# скрипт сам перепишет дату.
TEMPLATE_MIAC  = Path("МИАЦ_шаблон.docx")
TEMPLATE_FSTEK = Path("ФСТЭК_шаблон.docx")

# Куда складывать готовые файлы (папка создастся, если её нет).
OUTPUT_DIR = Path("отчёты")

# Префиксы для имён файлов. Итог: "ПРЕФИКС ДДММГГ.docx", например "МИАЦ 180526.docx"
PREFIX_MIAC  = "МИАЦ"
PREFIX_FSTEK = "ФСТЭК"

# Сколько дней генерировать (7 = понедельник–воскресенье).
DAYS_IN_WEEK = 7

# ============================================================================
# Дальше можно не редактировать
# ============================================================================

MONTHS_GENITIVE = {
    1: "января", 2: "февраля", 3: "марта",    4: "апреля",
    5: "мая",    6: "июня",    7: "июля",     8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}

# Фраза, которую ищем в склеенном тексте параграфа
_MONTHS_ALT = "|".join(MONTHS_GENITIVE.values())
PHRASE_RE = re.compile(
    rf"за\s+\d{{1,2}}\s+(?:{_MONTHS_ALT})\s+инцидентов\s+не\s+зарегистрировано",
    re.IGNORECASE,
)

# В docx XML текст параграфа может быть разорван на несколько <w:t> внутри разных <w:r>.
# Поэтому ищем целые параграфы и работаем с их содержимым.
PARAGRAPH_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.DOTALL)
WT_RE        = re.compile(r"<w:t(\s[^>]*)?>(.*?)</w:t>", re.DOTALL)


def build_phrase(d: date) -> str:
    """'за 18 мая инцидентов не зарегистрировано' и т.п."""
    return f"за {d.day} {MONTHS_GENITIVE[d.month]} инцидентов не зарегистрировано"


def _xml_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _replace_in_paragraph(p_xml, new_phrase):
    """
    Если параграф содержит целевую фразу — заменяет её на new_phrase,
    при этом сохраняя XML-структуру (форматирование первого run).

    Идея: собираем текст параграфа из всех <w:t>, ищем фразу в склейке,
    затем переписываем затронутые <w:t>. Всю новую фразу помещаем в первый
    затронутый <w:t> (где раньше начиналась старая фраза), а остальные
    затронутые <w:t> очищаем от текста, попавшего в диапазон.

    Возвращает (новый_xml_параграфа, количество_замен).
    """
    wt_matches = list(WT_RE.finditer(p_xml))
    if not wt_matches:
        return p_xml, 0

    # Карта сегментов: (старт_в_склейке, конец_в_склейке, match-объект для <w:t>)
    segments = []
    full_text = ""
    for wm in wt_matches:
        text = wm.group(2)
        segments.append((len(full_text), len(full_text) + len(text), wm))
        full_text += text

    pm = PHRASE_RE.search(full_text)
    if not pm:
        return p_xml, 0

    p_start, p_end = pm.start(), pm.end()

    # Затронутые сегменты — те, что пересекаются с диапазоном фразы
    affected = [(s, e, wm) for (s, e, wm) in segments if e > p_start and s < p_end]
    if not affected:
        return p_xml, 0

    # Перебираем затронутые с КОНЦА, чтобы смещения в p_xml не съезжали
    new_p = p_xml
    for seg_start, seg_end, wm in sorted(affected, key=lambda x: x[0], reverse=True):
        old_text = wm.group(2)
        before = old_text[: max(0, p_start - seg_start)] if seg_start < p_start else ""
        after  = old_text[ max(0, p_end   - seg_start):]  if seg_end > p_end else ""

        if seg_start <= p_start:
            # Сегмент содержит НАЧАЛО фразы — сюда складываем всю новую фразу
            content = before + new_phrase + after
        else:
            # Сегмент целиком внутри или содержит хвост — оставляем только хвост
            content = after

        attrs = wm.group(1) or ""
        if content:
            escaped = _xml_escape(content)
            # xml:space="preserve" нужен, если есть ведущие/завершающие пробелы
            if escaped != escaped.strip() and "xml:space" not in attrs:
                attrs += ' xml:space="preserve"'
            new_wt = f"<w:t{attrs}>{escaped}</w:t>"
        else:
            new_wt = f"<w:t{attrs}></w:t>"

        new_p = new_p[: wm.start()] + new_wt + new_p[wm.end():]

    return new_p, 1


def replace_in_docx(template, output, d):
    """
    Копирует template -> output, заменяя дату на d.
    Возвращает количество сделанных замен (должно быть >=1).
    """
    if not template.exists():
        raise FileNotFoundError(f"Шаблон не найден: {template}")

    new_phrase = build_phrase(d)

    with zipfile.ZipFile(template, "r") as zin:
        document_xml = zin.read("word/document.xml").decode("utf-8")

    new_xml = document_xml
    total = 0
    matches = list(PARAGRAPH_RE.finditer(document_xml))
    # Обходим параграфы с конца, чтобы смещения не съезжали
    for m in reversed(matches):
        new_p, n = _replace_in_paragraph(m.group(), new_phrase)
        if n:
            new_xml = new_xml[: m.start()] + new_p + new_xml[m.end():]
            total += n

    if total == 0:
        raise RuntimeError(
            f"В шаблоне {template} не найдена фраза "
            f"'за <число> <месяц> инцидентов не зарегистрировано'.\n"
            f"Открой шаблон в Word и убедись, что эта фраза присутствует."
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(template, "r") as zin, \
         zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            zout.writestr(item, data)

    return total


def find_monday(d):
    """Понедельник той недели, в которую попадает d."""
    return d - timedelta(days=d.weekday())


def parse_args():
    p = argparse.ArgumentParser(description="Генератор еженедельных отчётов МИАЦ/ФСТЭК")
    p.add_argument("start_date", nargs="?",
                   help="Дата в формате ДД.ММ.ГГГГ (по умолчанию — понедельник текущей недели)")
    p.add_argument("--next", action="store_true",
                   help="Следующая неделя относительно сегодня")
    args = p.parse_args()

    today = date.today()
    if args.start_date:
        try:
            day, month, year = map(int, args.start_date.split("."))
            start = date(year, month, day)
        except ValueError:
            sys.exit(f"Не понял дату '{args.start_date}'. Нужно ДД.ММ.ГГГГ, например 18.05.2026")
        return find_monday(start)
    if args.next:
        return find_monday(today) + timedelta(days=7)
    return find_monday(today)


def main():
    monday = parse_args()
    sunday = monday + timedelta(days=DAYS_IN_WEEK - 1)
    print(f"Генерирую отчёты с {monday.strftime('%d.%m.%Y')} по {sunday.strftime('%d.%m.%Y')}")
    print(f"Папка для вывода: {OUTPUT_DIR.resolve()}\n")

    day_names_ru = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    for i in range(DAYS_IN_WEEK):
        d = monday + timedelta(days=i)
        suffix = d.strftime("%d%m%y")  # ДДММГГ, например 180526
        print(f"{day_names_ru[d.weekday()]} {d.strftime('%d.%m.%Y')}:")
        replace_in_docx(TEMPLATE_MIAC,  OUTPUT_DIR / f"{PREFIX_MIAC} {suffix}.docx",  d)
        replace_in_docx(TEMPLATE_FSTEK, OUTPUT_DIR / f"{PREFIX_FSTEK} {suffix}.docx", d)
        print(f"  ✓ {PREFIX_MIAC} {suffix}.docx")
        print(f"  ✓ {PREFIX_FSTEK} {suffix}.docx")

    print(f"\nГотово. Файлов создано: {DAYS_IN_WEEK * 2}")


if __name__ == "__main__":
    main()
