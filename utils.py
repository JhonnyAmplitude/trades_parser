from typing import Any, Optional, Dict
from datetime import datetime

import xlrd

import re


from typing import Any, List, Optional, Tuple, Dict

import pandas as pd

from constants import CURRENCY_DICT, SPECIAL_OPERATION_HANDLERS, OPERATION_TYPE_MAP


def to_num(x: Any) -> float:
    """
    Преобразует вход в float, заменяя запятую на точку.
    Возвращает 0.0 при некорректном формате или NaN.
    """
    try:
        if pd.isna(x):
            return 0.0
        return float(str(x).replace(",", "."))
    except Exception:
        return 0.0


def extract_isin(comment: Any) -> str:
    """
    Ищет в тексте ISIN (12 символов, первые 2 — буквы).
    Работает даже если передали float/None.
    """
    text = str(comment or "")
    m = re.search(r"\b[A-Z]{2}[A-Z0-9]{10}\b", text)
    return m.group(0) if m else ""


def find_column_index(headers: List[str], *keywords: str) -> Optional[int]:
    """
    Находит индекс первой колонки, в названии которой есть все keywords.
    """
    for i, h in enumerate(headers):
        if all(k in h for k in keywords):
            return i
    return None


def find_header_row(df: pd.DataFrame, required: List[str]) -> Optional[int]:
    """
    Ищет в DataFrame первую строку, где встречаются все required keywords.
    Возвращает индекс строки или None.
    """
    for i, row in df.iterrows():
        cells = [str(c).lower() for c in row if pd.notna(c)]
        if all(any(req in cell for cell in cells) for req in required):
            return i
    return None


def find_block_start(df: pd.DataFrame, keyword: str) -> Optional[int]:
    """
    Находит первую строку, содержащую keyword, и возвращает следующий индекс.
    """
    for i, row in df.iterrows():
        text = " ".join(str(c).lower() for c in row if pd.notna(c))
        if keyword.lower() in text:
            return i + 1
    return None


def parse_header_data(row_str: str, header_data: Dict[str, Optional[str]]) -> None:
    """
    Извлекает из строки:
      - account_id       по шаблону "Генеральное соглашение: <число>"
      - account_date_start  по шаблону "от DD.MM.YYYY"
      - date_start, date_end по шаблону "Период: с DD.MM.YYYY по DD.MM.YYYY"
    """
    if "Генеральное соглашение:" in row_str:
        m = re.search(r"Генеральное соглашение:\s*(\d+)", row_str)
        if m:
            header_data["account_id"] = m.group(1)
        dm = re.search(r"от\s+(\d{2}\.\d{2}\.\d{4})", row_str)
        if dm:
            header_data["account_date_start"] = dm.group(1)

    elif "период:" in row_str.lower() and "по" in row_str.lower():
        parts = row_str.split()
        try:
            i_s = parts.index("с")
            i_p = parts.index("по")
            header_data["date_start"] = parts[i_s + 1]
            header_data["date_end"] = parts[i_p + 1]
        except (ValueError, IndexError):
            pass


def detect_operation_type(op: str, inc: Any, exp: Any) -> str:
    """
    Универсальный детектор типа операции.
    """
    if not isinstance(op, str):
        return "other"
    if "покупка" in op.lower():
        return "buy"
    if "продажа" in op.lower():
        return "sell"
    if op in SPECIAL_OPERATION_HANDLERS:
        return SPECIAL_OPERATION_HANDLERS[op](inc, exp)
    return OPERATION_TYPE_MAP.get(op, "other")




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

