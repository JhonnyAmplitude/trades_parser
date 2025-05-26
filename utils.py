from typing import Any, Optional, Dict
from datetime import datetime

import xlrd

import re


def parse_header_data(row_str: str, header_data: dict) -> None:
    """
    Извлекает из строки:
      - account_id       по шаблону "Генеральное соглашение: <число>"
      - account_date_start  по шаблону "от DD.MM.YYYY"
      - date_start, date_end по шаблону "Период: с DD.MM.YYYY по DD.MM.YYYY"
    """
    # Генеральное соглашение
    if "Генеральное соглашение:" in row_str:
        m = re.search(r"Генеральное соглашение:\s*(\d+)", row_str)
        if m:
            header_data["account_id"] = m.group(1)
        dm = re.search(r"от\s+(\d{2}\.\d{2}\.\d{4})", row_str)
        if dm:
            header_data["account_date_start"] = parse_date(dm.group(1))

    # Период
    elif "Период:" in row_str and "по" in row_str:
        # ожидаем вид: "Период:    с 01.07.2023 по 31.07.2023"
        parts = row_str.split()
        try:
            i_s = parts.index("с")
            i_p = parts.index("по")
            header_data["date_start"] = parse_date(parts[i_s + 1])
            header_data["date_end"]   = parse_date(parts[i_p + 1])
        except (ValueError, IndexError):
            pass

def is_nonzero(value: Any) -> bool:
    """
    Проверка на значение, отличное от нуля.
    """
    try:
        return float(str(value).replace(",", ".").replace(" ", "")) != 0
    except (ValueError, TypeError):
        return False


def parse_date(value: Any) -> Optional[str]:
    """
    Универсальный парсер даты.
    Поддерживает:
    - datetime.datetime
    - Excel float/int дату (как в .xls)
    - Строки в формате 'дд.мм.гггг' или 'дд.мм.гг'
    Возвращает строку в формате 'YYYY-MM-DD' или None.
    """
    if not value:
        return None

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (int, float)):
        try:
            date = datetime(*xlrd.xldate_as_tuple(value, 0))
            return date.strftime("%Y-%m-%d")
        except Exception:
            return None

    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

    return None

